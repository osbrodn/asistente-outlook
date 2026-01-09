/* global Office */

// CONFIGURACIÓN DE SEGURIDAD
const OPENAI_API_KEY = "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"; 
let synth = window.speechSynthesis;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Asignar eventos a los botones
        document.getElementById("btn-leer").onclick = () => procesarCorreo('leer');
        document.getElementById("btn-resumen").onclick = () => procesarCorreo('resumen');
        document.getElementById("btn-detener").onclick = () => detenerVoz();
    }
});

async function procesarCorreo(modo) {
    const output = document.getElementById("output");
    const btnLeer = document.getElementById("btn-leer");
    const btnResumen = document.getElementById("btn-resumen");

    // Bloquear botones durante el proceso
    setButtonsDisabled(true);
    output.innerText = "Accediendo al contenido del correo...";

    try {
        // 1. Obtener texto del correo
        const textoCorreo = await new Promise((resolve, reject) => {
            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
                else reject("Error al leer el correo.");
            });
        });

        if (modo === 'leer') {
            output.innerText = "Leyendo correo original...";
            hablar(textoCorreo);
        } else {
            output.innerText = "La IA está procesando y traduciendo...";
            const respuestaIA = await llamarIA(textoCorreo);
            output.innerText = respuestaIA;
            hablar(respuestaIA);
        }
    } catch (error) {
        output.innerText = "Error: " + error.message;
        console.error(error);
    } finally {
        setButtonsDisabled(false);
    }
}

async function llamarIA(texto) {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${OPENAI_API_KEY.trim()}`
        },
        body: JSON.stringify({
            model: "gpt-4o-mini",
            messages: [
                { 
                    role: "system", 
                    content: "Eres un asistente ejecutivo. Tu tarea es: 1. Si el correo no está en español, tradúcelo. 2. Resume el contenido en máximo 3 frases claras. 3. Usa un tono profesional en español de México." 
                },
                { role: "user", content: texto }
            ],
            max_tokens: 200
        })
    });

    const data = await response.json();
    if (!response.ok) throw new Error(data.error ? data.error.message : "Error en API");
    return data.choices[0].message.content;
}

function hablar(texto) {
    detenerVoz(); // Limpiar lecturas previas
    const utterance = new SpeechSynthesisUtterance(texto);
    
    // Configuración para México
    utterance.lang = 'es-MX';
    const voces = synth.getVoices();
    const vozMX = voces.find(v => v.lang.includes("MX") || v.name.includes("Mexico"));
    
    if (vozMX) utterance.voice = vozMX;
    utterance.rate = 1.0;
    
    synth.speak(utterance);
}

function detenerVoz() {
    if (synth.speaking) {
        synth.cancel();
    }
}

function setButtonsDisabled(state) {
    document.getElementById("btn-leer").disabled = state;
    document.getElementById("btn-resumen").disabled = state;
}