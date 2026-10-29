/*
 * Starts the editor's backend: node server/index.js. The settings are environment variables (see
 * server/app.js, server/auth.js and server/ldap.js); OSIGN_PORT is where it listens (8080).
 */
'use strict';
const http = require('node:http');
const { createApp } = require('./app');

const port = Number(process.env.OSIGN_PORT || 8080);
const server = http.createServer(createApp());
// A client that sends its request slowly holds a connection; not for long.
server.requestTimeout = 60000;
server.headersTimeout = 20000;
server.listen(port, () => console.info('NOBS OSign ' + require('../package.json').version + ' listening on port ' + port));
for (const sig of ['SIGTERM', 'SIGINT']) process.on(sig, () => server.close(() => process.exit(0)));
