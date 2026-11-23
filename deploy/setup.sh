#!/usr/bin/env bash
# Install / update the pdf2docx web UI (CPU only) on an Ubuntu server.
#
#   bash deploy/setup.sh                  # download model, build image, start the site
#   bash deploy/setup.sh --install-docker # same, and install Docker first if it is missing
#
# Safe to run again: it keeps .env, skips the model download when the weights are
# already there, rebuilds the image and restarts the container.
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")"
DEPLOY_DIR=$PWD
ROOT=$(dirname "$DEPLOY_DIR")
INSTALL_DOCKER=0
for arg in "$@"; do
  case $arg in
    --install-docker) INSTALL_DOCKER=1 ;;
    -h|--help) sed -n '2,8p' "$0"; exit 0 ;;
    *) echo "Unknown option: $arg" >&2; exit 2 ;;
  esac
done

say() { printf '\n==> %s\n' "$*"; }
die() { printf 'ERROR: %s\n' "$*" >&2; exit 1; }

# ---------------------------------------------------------------- Docker
if ! command -v docker >/dev/null 2>&1; then
  [ "$INSTALL_DOCKER" = 1 ] || die "Docker is not installed. Run again with --install-docker, or install it yourself: https://docs.docker.com/engine/install/ubuntu/"
  say "Installing Docker Engine from Docker's apt repository"
  sudo apt-get update
  sudo apt-get install -y ca-certificates curl
  sudo install -m 0755 -d /etc/apt/keyrings
  sudo curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
  sudo chmod a+r /etc/apt/keyrings/docker.asc
  . /etc/os-release
  echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu ${UBUNTU_CODENAME:-$VERSION_CODENAME} stable" \
    | sudo tee /etc/apt/sources.list.d/docker.list >/dev/null
  sudo apt-get update
  sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
fi

DOCKER="docker"
if ! docker info >/dev/null 2>&1; then
  sudo -n true 2>/dev/null || echo "Docker needs sudo for this user; you may be asked for your password."
  DOCKER="sudo docker"
  $DOCKER info >/dev/null 2>&1 || die "Cannot talk to the Docker daemon. Is it running?  (sudo systemctl start docker)"
fi
$DOCKER compose version >/dev/null 2>&1 || die "The Docker Compose plugin is missing:  sudo apt-get install docker-compose-plugin"

[ "$(uname -m)" = x86_64 ] || die "This setup supports x86_64 servers only (found $(uname -m))."

# ---------------------------------------------------------------- settings
if [ ! -f .env ]; then
  say "Creating deploy/.env"
  cp env.example .env
  password=$(head -c 18 /dev/urandom | base64 | tr -d '/+=' | cut -c1-20)
  sed -i "s|^WEB_PASSWORD=.*|WEB_PASSWORD=${password}|; s|^PUID=.*|PUID=$(id -u)|; s|^PGID=.*|PGID=$(id -g)|" .env
  chmod 600 .env
fi
get() { grep -E "^$1=" .env | tail -n1 | cut -d= -f2-; }
MODEL_REPO=$(get MODEL_REPO); MODEL_REVISION=$(get MODEL_REVISION)
PUID=$(get PUID); PGID=$(get PGID); BIND=$(get BIND); PORT=$(get PORT)
if [ -z "$(get WEB_PASSWORD)" ] && [ "$BIND" != "127.0.0.1" ]; then
  die "WEB_PASSWORD is empty but BIND=$BIND would expose the site without a login. Set a password in deploy/.env."
fi

# ---------------------------------------------------------------- disk and memory
free_gb=$(df -BG --output=avail "$ROOT" | tail -n1 | tr -dc '0-9')
[ "$free_gb" -ge 12 ] || die "Only ${free_gb} GB free in $ROOT; about 12 GB is needed (model 3 GB, image 3 GB, build space)."
mem_gb=$(awk '/MemTotal/ {printf "%d", $2/1024/1024}' /proc/meminfo)
[ "$mem_gb" -ge 6 ] || echo "WARNING: this server has ${mem_gb} GB of RAM; 8 GB or more is recommended."

# ---------------------------------------------------------------- model
mkdir -p "$ROOT/model" "$ROOT/webdata"
if [ -f "$ROOT/model/model.safetensors" ] && [ -f "$ROOT/model/pyproject.toml" ]; then
  say "Model already present in $ROOT/model"
else
  say "Downloading model ${MODEL_REPO}@${MODEL_REVISION:0:12} (about 2.7 GB)"
  $DOCKER run --rm --user "$(id -u):$(id -g)" -e HOME=/tmp -e PYTHONUSERBASE=/tmp/py \
    -v "$ROOT/model:/model" python:3.12-slim sh -c '
      pip install --quiet --user --no-warn-script-location "huggingface_hub>=1.24,<2" &&
      python -c "
import sys
from huggingface_hub import snapshot_download
snapshot_download(sys.argv[1], revision=sys.argv[2] or None, local_dir=\"/model\")
" "$0" "$1"' "$MODEL_REPO" "$MODEL_REVISION"
fi
say "Checking model checksum"
(cd "$ROOT/model" && sha256sum -c SHA256SUMS) || die "The model file is corrupted. Delete $ROOT/model/model.safetensors and run this script again."

# webdata must be writable by the container user
if [ "$(stat -c %u "$ROOT/webdata")" != "$PUID" ]; then
  sudo chown -R "$PUID:$PGID" "$ROOT/webdata"
fi

# ---------------------------------------------------------------- build and start
say "Building the image (first time: 5-15 minutes)"
$DOCKER compose build
say "Starting the site"
$DOCKER compose up -d

printf 'Waiting for the model to load'
for _ in $(seq 1 90); do
  state=$($DOCKER inspect -f '{{.State.Status}}' pdf2docx-web 2>/dev/null || echo missing)
  [ "$state" = running ] || { echo; $DOCKER compose logs --tail 40; die "The container is not running (state: $state)."; }
  if $DOCKER exec pdf2docx-web curl -fsS http://127.0.0.1:8080/healthz 2>/dev/null | grep -q '"ready"'; then
    echo " ready"; ready=1; break
  fi
  printf '.'; sleep 3
done
[ "${ready:-0}" = 1 ] || { echo; $DOCKER compose logs --tail 40; die "The model did not finish loading in time."; }

host=$BIND; [ "$BIND" = 0.0.0.0 ] && host=$(hostname -I 2>/dev/null | awk '{print $1}')
cat <<EOF

Done.
  Address : http://${host:-127.0.0.1}:${PORT}
  User    : $(get WEB_USER)
  Password: $(get WEB_PASSWORD)   (stored in deploy/.env)

  Logs    : cd $DEPLOY_DIR && $DOCKER compose logs -f
  Stop    : cd $DEPLOY_DIR && $DOCKER compose down
  Start   : cd $DEPLOY_DIR && $DOCKER compose up -d
EOF
if [ "$BIND" = 127.0.0.1 ]; then
  cat <<EOF

The site only listens on this server (BIND=127.0.0.1). To use it from your PC either
  - open an SSH tunnel:  ssh -L ${PORT}:127.0.0.1:${PORT} <user>@<server>   then browse http://localhost:${PORT}
  - or set BIND=0.0.0.0 in deploy/.env and run:  $DOCKER compose up -d
EOF
fi
