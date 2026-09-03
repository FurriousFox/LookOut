import { getXOAUTH2 } from "./src/auth.ts";
import tls from "node:tls";
let d = "";

console.time("obtaining token");
const token = await getXOAUTH2();
console.timeEnd("obtaining token");

const socket = tls.connect({
    host: "outlook.office365.com",
    port: 993,
    servername: "outlook.office365.com",
}, async () => {
    console.log("connection established");

    socket.write(`0 CAPABILITY\r\n`);
    socket.write(`1 AUTHENTICATE XOAUTH2 ${token}\r\n`);
    let i = -100;
    while (++i) {
        if (d.includes("\r\n1 OK")) break;
        await new Promise(r => setTimeout(r, 10));
    }
    if (i) console.log("authentication success");
    else console.log("authentication failure");

    socket.write(`2 LOGOUT\r\n`);
});
socket.on('data', data => d += data);
socket.on('end', () => console.log('connection ended'));
