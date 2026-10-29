# OSign, the editor for the signatures, as a server (server/), for Docker Compose behind the shared
# proxy (see docker-compose.yml and deploy-osign.sh). The add-in itself is not served from here.

# The add-in's template (dist/template) is built with the dev tools, which the image leaves out.
FROM node:26.12-alpine3.24 AS build
WORKDIR /src
COPY package.json package-lock.json ./
RUN npm ci --ignore-scripts
COPY addin addin
COPY shared shared
RUN node addin/build.mjs

FROM node:26.12-alpine3.24
# Security fixes Alpine has published since the base image was built (OpenSSL, ...).
RUN apk upgrade --no-cache
WORKDIR /app
ENV NODE_ENV=production OSIGN_DATA=/data OSIGN_PORT=8080 TZ=Europe/Zurich
COPY package.json package-lock.json ./
# npm is only needed to install. It is removed afterwards, so the dependencies it bundles are not in the image.
RUN npm ci --omit=dev --ignore-scripts && npm cache clean --force \
 && rm -rf /usr/local/lib/node_modules/npm /usr/local/lib/node_modules/corepack \
           /usr/local/bin/npm /usr/local/bin/npx /usr/local/bin/corepack /root/.npm
COPY server ./server
COPY shared ./shared
COPY editor/ui ./editor/ui
COPY addin/fill.mjs ./addin/fill.mjs
COPY --from=build /src/dist/template ./addin-template
RUN mkdir -p /data && chown node:node /data
USER node
VOLUME ["/data"]
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=3s CMD wget -qO- http://127.0.0.1:8080/health > /dev/null || exit 1
CMD ["node", "server/index.js"]
