/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
const keypress = require("keypress");
const WebSocketServer = require("ws").WebSocketServer;

const wssConfig = { port: 9000 };
const wss = new WebSocketServer(wssConfig);

let lastConnection;

wss.on("connection", function connection(ws) {
    console.log("connected")

    ws.on("message", function message(data) {
        console.log("received: %s", data);
    });

    lastConnection = ws;
});

keypress(process.stdin);

process.stdin.on("keypress", function (ch, key) {
    console.log(`got "keypress"`, key);

    if (key && key.ctrl && key.name == "c") {
        process.stdin.pause();
        process.exit(0);
    }

    if (key.name == "n") {
        lastConnection.send("new_message");
    }

    if (key.name == "r") {
        lastConnection.send("reply_to");
    }

    if (key.name == "a") {
        lastConnection.send("reply_to_all");
    }

    if (key.name == "f") {
        lastConnection.send("forward_to");
    }

    if (key.name == "t") {
        lastConnection.send("tag_message");
    }

    if (key.name == "u") {
        lastConnection.send("untag_message");
    }

    if (key.name == "c") {
        lastConnection.send("create_foulder");
    }

    if (key.name == "d") {
        lastConnection.send("download_attachments");
    }
});

process.stdin.setRawMode(true);
process.stdin.resume();