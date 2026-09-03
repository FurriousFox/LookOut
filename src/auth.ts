import process from "node:process";
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

import type { INetworkModule } from "@azure/msal-common";
import * as msal from "@azure/msal-node";
import puppeteer from "puppeteer";
import type * as Puppeteer from "puppeteer";

const CLIENT_ID = "5801fbf1-8df2-4c3c-a65c-e230edf60ca2"; // Microsoft Remote Connectivity Analyzer
const REDIRECT_URI = "https://testconnectivity.microsoft.com/authcallback";
const SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "openid",
    "profile",
    "offline_access",
];

if (!fs.existsSync(path.dirname(`${import.meta.dirname}/../state/`))) fs.mkdirSync(path.dirname(`${import.meta.dirname}/../state/`), { recursive: true });
const MSAL_CREDS = `${import.meta.dirname}/../state/msal_creds.json`;

const spoofClient: INetworkModule = {
    sendGetRequestAsync: async (url, options) => {
        const res = await fetch(url, { method: "GET", headers: { ...options?.headers, Origin: "https://testconnectivity.microsoft.com" } });
        return { headers: Object.fromEntries(res.headers.entries()), body: await res.json(), status: res.status };
    },
    sendPostRequestAsync: async (url, options) => {
        const res = await fetch(url, { method: "POST", headers: { ...options?.headers, Origin: "https://testconnectivity.microsoft.com" }, body: options?.body });
        return { headers: Object.fromEntries(res.headers.entries()), body: await res.json(), status: res.status };
    }
};

const pca = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/common` },
    system: { networkClient: spoofClient }
});

export async function getToken(forceInteractive = false, silent = false, tryHeadless = false, returnUrl = false, customMsalUrl: undefined | string = undefined) {
    let needsInteractive = true;

    try {
        if (fs.existsSync(MSAL_CREDS)) {
            const data = fs.readFileSync(MSAL_CREDS, "utf8");
            if (data && data.length > 0) {
                pca.getTokenCache().deserialize(data);
                needsInteractive = false;
            }
        } else if (!forceInteractive && fs.existsSync(path.dirname(`${import.meta.dirname}/../state/puppeteerUserData/`))) {
            tryHeadless = true;
        }
    } catch (err) {
        console.log("MSAL credentials data invalid:", err);
    }

    if (needsInteractive || forceInteractive) {
        const codeVerifier = crypto.randomBytes(32).toString("base64url");
        const codeChallenge = crypto.createHash("sha256")
            .update(codeVerifier)
            .digest()
            .toString("base64url");

        const code_url: Promise<string> = new Promise((resolve, reject) => (async (resolve, reject) => {
            let a = false;
            const msal_url = customMsalUrl ? customMsalUrl : await pca.getAuthCodeUrl({
                scopes: SCOPES,
                redirectUri: REDIRECT_URI,
                codeChallenge: codeChallenge,
                codeChallengeMethod: "S256"
            });

            let f = false;
            let g: number | undefined | NodeJS.Timeout;
            const browser_obj = {
                browser: await puppeteer.launch({
                    ...tryHeadless ? {
                        headless: true,
                    } : {
                        headless: false,
                        defaultViewport: null,
                    },
                    userDataDir: `${import.meta.dirname}/../state/puppeteerUserData`,
                    args: [
                        "--no-sandbox",
                        '--no-first-run',
                        '--no-default-browser-check',
                        '--disable-session-crashed-bubble',
                        '--disable-restore-session-state',
                        '--disable-gpu',
                        '--disable-software-rasterizer',
                        '--disable-features=Translate',
                    ],
                })
            };
            const browserWsEndpoint = browser_obj.browser.wsEndpoint();

            if (tryHeadless) {
                if (!silent) console.log("trying to obtain token semi-interactively (headless), timeout: 15 seconds...");
                g = setTimeout(async () => {
                    if (f) return; else f = true;
                    if (!silent) console.log("failed to headlessly obtain token, user interaction required");

                    browser_obj.browser.close();
                    resolve(await getToken(true, undefined, false, true, msal_url));
                }, 15000);
            }

            process.on("SIGINT", () => {
                browser_obj.browser.close(); f = true;
                reject("received SIGINT");

                process.exit();
            });
            browser_obj.browser.on("disconnected", async () => {
                const new_browser = await puppeteer.connect({ browserWSEndpoint: browserWsEndpoint, defaultViewport: null }).catch(() => { if (!a && !f) { f = true; reject("browser closed"); } });
                if (new_browser) {
                    browser_obj.browser = new_browser; (pageobject.page = (await browser_obj.browser.pages())[0]);
                    registerFramenavigated(pageobject as { page: Puppeteer.Page; });
                } else {
                    if (!a && !f) { f = true; reject("browser closed"); }
                }
            });

            const pageobject: { page?: Puppeteer.Page; } = {};
            if ((await browser_obj.browser.pages()).length) {
                (pageobject.page = (await browser_obj.browser.pages())[0]);
            } else {
                (pageobject.page = await browser_obj.browser.newPage());
            }
            await pageobject.page.bringToFront();
            await pageobject.page.goto("about:blank");
            try {
                const pages = await browser_obj.browser.pages();
                while (pages.length > 1) await pages.pop()?.close({ runBeforeUnload: false });
            } catch (_e) { /*  */ }

            function registerFramenavigated(pageobject: { page: Puppeteer.Page; }) {
                pageobject.page.on("framenavigated", async (frame) => {
                    const url = frame.url();
                    if (url.match(/https:\/\/login\.microsoftonline\.com\/.+\/.+/))
                        pageobject.page.evaluate(`{ const rememberYes = setInterval(()=>{ if (document.querySelector(".button_primary") && document.querySelector("meta[name=PageID]").content == 'KmsiInterrupt') { clearInterval(rememberYes); document.querySelector(".button_primary").click(); } }, 100) }`).catch(() => { });
                    else if (url.match(/https:\/\/login\.live\.com\/.+\.srf/))
                        pageobject.page.evaluate(`{ const rememberYes = setInterval(()=>{ if (document.querySelector("button[data-testid=primaryButton]") && !!(document.querySelector("img[data-testid=kmsiImage]")??document.querySelector("div[data-testid=kmsiVideo]"))) { clearInterval(rememberYes); document.querySelector("button[data-testid=primaryButton]").click(); } }, 100) }`).catch(() => { });

                    if (url.startsWith(REDIRECT_URI) && !a) {
                        a = true;
                        await browser_obj.browser.close();
                        f = true;
                        if (g) clearTimeout(g as (number | undefined));
                        resolve(url);
                    }
                });
            }

            registerFramenavigated(pageobject as { page: Puppeteer.Page; });

            pageobject.page.goto(msal_url).catch(_e => { /*  */ });
        })(resolve, reject));


        if (returnUrl) return await code_url;

        const code = new URL(await code_url).searchParams.get("code");
        if (!code) throw new Error("Failed to interactively obtain token");

        const token = await pca.acquireTokenByCode({
            code,
            scopes: SCOPES,
            redirectUri: REDIRECT_URI,
            codeVerifier: codeVerifier
        });
        if (!token.accessToken) throw new Error("no access token received");

        fs.writeFileSync(MSAL_CREDS, pca.getTokenCache().serialize());


        if (!silent) console.log("interactively obtained access token, expires:", token.expiresOn);
        return token.accessToken;
    } else {
        const accounts = await pca.getTokenCache().getAllAccounts();
        if (accounts.length === 0) return await getToken(true);

        try {
            const token = await pca.acquireTokenSilent({
                account: accounts[0],
                scopes: SCOPES,
            });

            fs.writeFileSync(MSAL_CREDS, pca.getTokenCache().serialize());

            if (!silent) console.log(`silently obtained access token, expires:`, token.extExpiresOn, token.account?.username);
            return token.accessToken;
        } catch (_err) {
            console.log(/* _err, "\n */"failed to silently obtain token, trying headless browser");
            return await getToken(true, undefined, true);
        }
    }
}

export async function getXOAUTH2() {
    await getToken();

    try {
        const data = fs.readFileSync(MSAL_CREDS, "utf8");
        if (data && data.length > 0) {
            pca.getTokenCache().deserialize(data);
        }
    } catch (err) { throw new Error(`Failed to import MSAL credentials: ${(err as Error).stack}`); }

    const accounts = await pca.getTokenCache().getAllAccounts();
    const token = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: SCOPES,
    });

    if (token.account?.username == undefined) throw new Error("account username is undefined");
    return btoa(`user=${token.account.username}\x01auth=Bearer ${token.accessToken}\x01\x01`);
}

export default getToken;