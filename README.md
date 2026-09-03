# LookOut

IMAP proxy imitating [a Microsoft tool](https://testconnectivity.microsoft.com/tests/O365Imap/input) to provide an IMAP proxy for Microsoft Exchange, bypassing the need for an organization admin to allow your OAuth2 app.

## config.json

```jsonc
{
    // IMAP proxy port
    "port": 993,
    // 'none' / SSL encryption, STARTTLS is unsupported
    "ssl": true,
    
    // cert chain in PEM format
    "certchain": "/etc/letsencrypt/live/example.com/fullchain.pem",
    // private key in PEM format
    "certkey": "/etc/letsencrypt/live/example.com/privkey.pem",

    // sha512 hash of the password used for authentication
    // only the 'LOGIN' command is supported for auth
    "password_sha512": "3bb12eda3c298db5de25597f54d924f2e17e78a26ad8953ed8218ee682f0bbbe9021e2f3009d152c911bf1f25ec683a902714166767afbd8e5bd0fb0124ecb8a",

    // whether to run a fake smtp server (no encryption)
    // doesn't deliver mail, only passes simple connection tests
    "fake_smtp": true,
    // fake smtp server port
    "fake_smtp_port": 587
}
```

## usage

```sh
node test.ts #run a connection test before usage, recommended
node server.ts #run server, state is stored in ./state/
```

---

**note:** NodeJS is strongly recommended over Deno/Bun because of stability issues with Puppeteer
