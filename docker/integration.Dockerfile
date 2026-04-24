# Integration-test image: Thunderbird + Node + test tooling, ready to drive
# the MCP bridge against a live TB instance talking to a Greenmail IMAP mock.
#
# Pinned to a specific Thunderbird release so the test surface is
# reproducible -- Ubuntu's apt package lags and a floating "latest"
# breaks the next time Mozilla ships a breaking change. Bump when ready.
#
# Build locally:
#   docker build -f docker/integration.Dockerfile -t thunderbird-mcp-ci .
# Run locally:
#   docker run --rm --network host -v $(pwd):/work -w /work thunderbird-mcp-ci \
#     bash -c "scripts/ci/make-tb-profile.sh /tmp/profile && test/integration/run.sh"

FROM ubuntu:24.04

# Thunderbird ESR tarball pin. Matches strict_max_version="150.*" in our
# extension/manifest.json. Change ESR_VERSION to rebuild against a
# different release; the tarball URL scheme is stable across versions.
ARG ESR_VERSION=128.18.0esr
ARG ESR_LOCALE=en-US

ENV DEBIAN_FRONTEND=noninteractive
ENV NODE_VERSION=22
ENV TB_HOME=/opt/thunderbird
ENV PATH=${TB_HOME}:${PATH}

RUN apt-get update && apt-get install -y --no-install-recommends \
      ca-certificates curl bzip2 xz-utils \
      # Thunderbird's runtime deps (even in --headless TB loads gtk/dbus)
      libgtk-3-0 libasound2t64 libx11-xcb1 libxcomposite1 libxdamage1 \
      libxext6 libxfixes3 libxrandr2 libgbm1 libpango-1.0-0 libpangocairo-1.0-0 \
      libnss3 libnspr4 libcups2 libdrm2 libxkbcommon0 libatk1.0-0 \
      libatk-bridge2.0-0 libatspi2.0-0 libdbus-glib-1-2 libxtst6 \
      # For createDrafts / IMAP sync we need a dbus session
      dbus dbus-x11 \
      # netcat for readiness loops; jq for JSON munging in shell scripts
      netcat-openbsd jq \
      && rm -rf /var/lib/apt/lists/*

# Node. Uses NodeSource so we get a recent version instead of Ubuntu's stale one.
RUN curl -fsSL https://deb.nodesource.com/setup_${NODE_VERSION}.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

# Thunderbird tarball (Mozilla release CDN, no apt dependency)
RUN curl -fsSL \
      "https://ftp.mozilla.org/pub/thunderbird/releases/${ESR_VERSION}/linux-x86_64/${ESR_LOCALE}/thunderbird-${ESR_VERSION}.tar.xz" \
      -o /tmp/thunderbird.tar.xz \
    && mkdir -p "${TB_HOME}" \
    && tar -xJf /tmp/thunderbird.tar.xz -C "${TB_HOME}" --strip-components=1 \
    && rm /tmp/thunderbird.tar.xz \
    && ln -s "${TB_HOME}/thunderbird" /usr/local/bin/thunderbird \
    && thunderbird --version

WORKDIR /work

# Sanity: verify TB can print its version in headless mode inside the
# container. Fails fast at build time if the runtime deps above are wrong.
RUN thunderbird --headless --version

CMD ["bash"]
