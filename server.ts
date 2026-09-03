import tls from "node:tls";
import net from "node:net";
import process from "node:process";
import fs from "node:fs";
import readline from "node:readline/promises";
import crypto from "node:crypto";
import { getXOAUTH2 } from "./src/auth.ts";
import _config from "./config.json" with {type: "json"};
function Tag(l = 8) {
    return Array.from(
        crypto.getRandomValues(new Uint32Array(l)),
        value => 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'[value % 52]
    ).join('');
}

type config = {
    port: number;
    ssl: boolean;
    certchain?: string;
    certkey?: string;
    password_sha512: string;
    fake_smtp: boolean;
    fake_smtp_port?: number;
};
const config: config = _config;

const verbose = false;

process.on("uncaughtException", e => console.log("uncaughtException", e, `\n${e.stack}`));
process.on("unhandledRejection", e => console.log("unhandledRejection", e));

const server: tls.Server | net.Server = config.ssl ?
    tls.createServer({
        cert: config.certchain ? fs.readFileSync(config.certchain) : undefined,
        key: fs.readFileSync(config.certkey!),
    }, socketHandler)
    : net.createServer(socketHandler);

async function socketHandler(socket: net.Socket | tls.TLSSocket) {
    socket.on('error', () => { try { client.destroy(); } catch (_e) { /*  */ } });
    socket.on('end', () => { try { client.end(); } catch (_e) { /*  */ } });

    if (verbose) console.log("SERVER:", "* OK Microsoft Exchange Something Something");
    socket.write("* OK Microsoft Exchange Something Something\r\n");
    let tag = "0";
    let auth = false;

    const sr = readline.createInterface({
        input: socket,
        crlfDelay: Infinity,
        tabSize: 1,
    });
    const lines: string[] = [];
    sr.on('line', line => {
        lines.push(line);
        if (verbose) console.log("CLIENT:", line);
    });

    loop: while (true) {
        let i = 0;
        if (lines.length <= 0) {
            await new Promise(r => setTimeout(r, 10));
            if (i++ > 1000) throw new Error("timeout exceeded (10s)");
            continue loop;
        }

        const line = lines.shift()!;
        const match = line.match(/^(?<tag>[A-Za-z0-9][A-Za-z0-9._-]*)[ \t]+CAPABILITY[ \t]*$/i);
        if (match) {
            tag = match[1];
            break;
        }
    }

    let r: (value: void | PromiseLike<void>) => void;
    const p = new Promise<void>(e => r = e);
    const client = tls.connect({
        host: "outlook.office365.com",
        port: 993,
        servername: "outlook.office365.com",
    }, r!);
    client.on('error', () => socket.destroy());
    client.on('end', () => socket.end());
    {
        let i = 0;
        const sr = readline.createInterface({
            input: client,
            crlfDelay: Infinity,
            tabSize: 1,
        });
        const lines: string[] = [];
        sr.on('line', line => {
            lines.push(line);
            if (verbose) console.log("SERVER:", line);
        });
        await p;

        const tag = Tag();
        client.write(`${tag} CAPABILITY\r\n`);
        loop: while (true) {
            if (lines.length <= 0) {
                await new Promise(r => setTimeout(r, 10));
                if (i++ > 1000) throw new Error("timeout exceeded (10s)");
                continue loop;
            }
            const line = lines.shift()!;
            const match = line.match(/^\*[ \t]+CAPABILITY[ \t]+(.*)$/i);
            if (match) {
                const capabilities = match[1]
                    .replaceAll(/[ \t]*LOGINDISABLED/gi, '')
                    .replaceAll(/[ \t]*AUTH=[^ \t]*/gi, '');
                if (verbose) console.log("SERVER:", `* CAPABILITY ${capabilities}`);
                socket.write(`* CAPABILITY ${capabilities}\r\n`);
                break loop;
            }
        }
        loop: while (true) {
            if (lines.length <= 0) {
                await new Promise(r => setTimeout(r, 10));
                if (i++ > 1000) throw new Error("timeout exceeded (10s)");
                continue loop;
            }
            const line = lines.shift()!;
            if (line.startsWith(`${tag}\t`) || line.startsWith(`${tag} `)) {
                break loop;
            }
        }
    }
    if (verbose) console.log("SERVER:", `${tag} OK CAPABILITY completed.`);
    socket.write(`${tag} OK CAPABILITY completed.\r\n`);

    sr.on('line', async line => {
        if (auth !== false) return;

        // deno-lint-ignore no-control-regex
        const match = line.match(/^(?<tag>[A-Za-z0-9][A-Za-z0-9._-]*) LOGIN (?:"(?<username>(?:[^"\\\r\n]|\\[^\r\n])*)"|(?<username>[^\x00-\x1F\x7F (){%*"\\\r\n]+)) (?:"(?<password>(?:[^"\\\r\n]|\\[^\r\n])*)"|(?<password>[^\x00-\x1F\x7F (){%*"\\\r\n]+))$/i);
        if (match && match.groups && match.groups.username && match.groups.password && match.groups.tag) {
            // if (match.groups.username.toLowerCase() != "john@example.com") return;
            const password = crypto.createHash('sha512').update(match.groups.password).digest('hex');
            if (password == config.password_sha512) {
                console.log("login correct");
                auth = true;
            }
            else { sr.close(); client.destroy(); socket.destroy(); return; };
        } else return;

        if (auth !== true) throw new Error("help, this is very wrongggg!");

        if (verbose) console.log("\n\n\n");

        socket.pipe(client);
        client.pipe(socket);

        client.write(`${match!.groups!.tag} AUTHENTICATE XOAUTH2 ${await getXOAUTH2()}\r\n`);
    });
}

server.listen(config.port, () => console.log(`listening on port ${config.port}`));

if (config.fake_smtp) {
    net.createServer(socket => {
        socket.on("error", () => { /*  */ });
        socket.write("220 fakesmtp ESMTP ready\r\n");
        const sr = readline.createInterface({
            input: socket,
            crlfDelay: Infinity,
            tabSize: 1,
        });
        sr.on('line', line => {
            if (line.toUpperCase().startsWith("EHLO ")) {
                socket.write("250-fakesmtp\r\n250 AUTH PLAIN\r\n");
            }
            if (line.toUpperCase().startsWith("AUTH PLAIN")) {
                socket.write("235 2.7.0 Authentication successful\r\n");
            }
            if (line.toUpperCase().startsWith("QUIT")) {
                socket.write("221 2.0.0 Bye\r\n");
                socket.end();
            }
        });
    }).listen(config.fake_smtp_port, () => console.log(`listening on port ${config.fake_smtp_port} (fake smtp)`));
}
