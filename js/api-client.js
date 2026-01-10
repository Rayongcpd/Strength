// Configuration
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzMW3yBvzK31hXgwFhfkmK0SvEgRrPnpL5WuBDcrw4Rip9wL0iM9p4m_boZ_Z7aKdjA/exec"; // <--- ใส่ Web App URL ของคุณที่นี่

/**
 * Google Apps Script API Client Helper
 * Emulates google.script.run for GitHub Pages
 */
class GoogleScriptRun {
    constructor() {
        this._successHandler = null;
        this._failureHandler = null;
    }

    withSuccessHandler(func) {
        this._successHandler = func;
        return this;
    }

    withFailureHandler(func) {
        this._failureHandler = func;
        return this;
    }

    // Generic handler using Proxy to capture any function name
}

const runnerHandler = {
    get: function (target, prop) {
        if (prop === 'withSuccessHandler' || prop === 'withFailureHandler') {
            return function (func) {
                target[prop](func);
                return new Proxy(target, runnerHandler);
            };
        }

        // Return a function that performs the API call
        return function (...args) {
            if (WEB_APP_URL === "YOUR_WEB_APP_URL_HERE") {
                const msg = "กรุณาใส่ Web App URL ในไฟล์ js/api-client.js";
                console.error(msg);
                if (target._failureHandler) target._failureHandler(new Error(msg));
                else alert(msg);
                return;
            }

            console.log(`Calling API: ${prop}`, args);

            fetch(WEB_APP_URL, {
                method: "POST",
                // Use text/plain to avoid CORS preflight
                headers: { "Content-Type": "text/plain;charset=utf-8" },
                body: JSON.stringify({
                    functionName: prop,
                    parameters: args
                })
            })
                .then(response => response.json())
                .then(data => {
                    // Expecting backend to return the direct result or standard envelope?
                    // Let's assume backend returns pure result to match google.script.run behavior
                    // BUT, if error occurs in GAS, it often returns HTML error page if not caught.
                    // We will implement safe execution in backend.

                    if (data && data.error) {
                        // Custom error protocol
                        if (target._failureHandler) target._failureHandler(new Error(data.error));
                    } else {
                        if (target._successHandler) target._successHandler(data);
                    }
                })
                .catch(err => {
                    console.error("API Error:", err);
                    if (target._failureHandler) target._failureHandler(err);
                });
        };
    }
};

window.google = {
    script: {
        run: new Proxy(new GoogleScriptRun(), runnerHandler)
    }
};
