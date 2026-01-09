/* global Office */

const OPENAI_API_KEY = "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"; 

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        const btn = document.getElementById("run-ai");
        if (btn) btn.onclick = iniciarResumen;
    }
});

async function iniciarResumen() {
    const btn = document.getElementById("run-ai");
    const status = document.getElementById("loading-text");
    const output = document.getElementById("output");

    btn.disabled = true;
    status.style.display = "block";
    output.innerText = "Leyendo correo...";

    try {
        const mailBody = await new Promise((resolve, reject) => {
            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
                else reject("No se pudo leer el correo.");
            });
        });

        output.innerText = "Consultando a la IA...";

        const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${OPENAI_API_KEY.trim()}`
            },
            body: JSON.stringify({
                model: "gpt-4o-mini",
                messages: [
                    { role: "system", content: "Eres un asistente ejecutivo. Resume en 2 frases máximo en español de México." },
                    { role: "user", content: mailBody }
                ],
                max_tokens: 150
            })
        });

        const data = await response.json();
        if (!response.ok) throw new Error(data.error ? data.error.message : "Error en OpenAI");

        const resumen = data.choices[0].message.content;
        output.innerText = resumen;

        const msg = new SpeechSynthesisUtterance(resumen);
        msg.lang = "es-MX";
        window.speechSynthesis.speak(msg);

    } catch (error) {
        output.innerText = "Error: " + error.message;
    } finally {
        btn.disabled = false;
        status.style.display = "none";
    }
}