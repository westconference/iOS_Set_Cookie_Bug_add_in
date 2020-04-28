const express = require('express')
const app = express()
const port = process.env.addin_service_port || 3000;
const https = require('https');
const fs = require('fs');
const path = require('path');

app.use('/assets', express.static('assets'));

app.get('/commands.html', (req, res) => {
	res.sendFile(path.join(__dirname + '/src/commands/commands.html'));
})
app.get('/commands.js', (req, res) => {
	res.sendFile(path.join(__dirname + '/src/commands/commands.js'));
})

app.get('/dialog.html', (req, res) => {
	res.sendFile(path.join(__dirname + '/src/dialog/dialog.html'));
})

app.get('/return.html', (req, res) => {
	res.sendFile(path.join(__dirname + '/src/dialog/return.html'));
})

app.get('/manifest.xml', (req, res) => {
	res.sendFile(path.join(__dirname + '/manifest.xml'));
})

https.createServer({
  key: fs.readFileSync('certs/server.key'),
  cert: fs.readFileSync('certs/server.cert')
}, app).listen(port, () => {
  console.log('Listening on port ' + port)
})