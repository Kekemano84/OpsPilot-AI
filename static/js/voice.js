function startTrailerVoice() {
    const status = document.getElementById("voiceStatus");
    const trailerInput = document.getElementById("voiceTrailer");
    const locationType = document.getElementById("voiceLocationType");
    const locationDetail = document.getElementById("voiceLocationDetail");
    const notes = document.getElementById("voiceNotes");
    const source = document.getElementById("voiceSource");

    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
        if (status) status.innerText = "Voice input is not supported in this browser. Try Chrome.";
        return;
    }

    const recognition = new SpeechRecognition();
    recognition.lang = "en-GB";
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    if (status) status.innerText = "Listening...";
    recognition.start();

    recognition.onresult = function(event) {
        const text = event.results[0][0].transcript;
        if (status) status.innerText = "Heard: " + text;
        if (source) source.value = "Voice";

        const upper = text.toUpperCase();
        const trailerMatch = upper.match(/[A-Z]{1,4}\s?\d{2,6}/);
        if (trailerMatch && trailerInput) trailerInput.value = trailerMatch[0].replace(/\s+/g, "");

        const doorMatch = upper.match(/DOOR\s?\d+/);
        const fenceMatch = upper.match(/FENCE\s?[A-Z0-9]+/);
        const yardMatch = upper.match(/YARD\s?[A-Z0-9]+/);

        if (doorMatch) {
            locationType.value = "Door";
            if (typeof toggleYardLocationFields === "function") toggleYardLocationFields();
            const doorNumber = document.getElementById("doorNumber");
            if (doorNumber) doorNumber.value = doorMatch[0].replace(/DOOR\s?/i, "Door ");
            if (locationDetail) locationDetail.value = doorMatch[0];
        } else if (fenceMatch) {
            locationType.value = "Fence";
            if (typeof toggleYardLocationFields === "function") toggleYardLocationFields();
            const fenceNumber = document.getElementById("fenceNumber");
            if (fenceNumber) fenceNumber.value = fenceMatch[0].replace(/FENCE\s?/i, "Fence ");
            if (locationDetail) locationDetail.value = fenceMatch[0];
        } else if (yardMatch) {
            locationType.value = "Yard";
            if (typeof toggleYardLocationFields === "function") toggleYardLocationFields();
            if (locationDetail) locationDetail.value = yardMatch[0];
        }

        if (notes) notes.value = text;
    };

    recognition.onerror = function(event) {
        if (status) status.innerText = "Voice error: " + event.error;
    };
}


function parseBulkTrailerSpeech(text) {
    let normalized = text
        .replace(/\band\b/gi, ",")
        .replace(/\bthen\b/gi, ",")
        .replace(/\bnext\b/gi, ",")
        .replace(/\./g, ",")
        .replace(/\s+/g, " ")
        .trim();

    const chunks = normalized.split(",").map(x => x.trim()).filter(Boolean);
    const lines = [];

    for (const chunk of chunks) {
        const trailerMatch = chunk.match(/\b[A-Z]{1,5}\s?\d{2,6}\b/i);
        const doorMatch = chunk.match(/\bdoor\s?(\d{1,4})\b/i);
        const fenceMatch = chunk.match(/\bfence\s?(\d{1,4})\b/i);

        if (!trailerMatch) continue;

        const trailer = trailerMatch[0].replace(/\s+/g, "").toUpperCase();
        let location = "Yard";

        if (doorMatch) location = "Door " + doorMatch[1];
        if (fenceMatch) location = "Fence " + fenceMatch[1];

        let note = chunk
            .replace(trailerMatch[0], "")
            .replace(doorMatch ? doorMatch[0] : "", "")
            .replace(fenceMatch ? fenceMatch[0] : "", "")
            .trim();

        if (!note) note = "Recorded by voice";

        lines.push(`${trailer}, ${location}, ${note}`);
    }

    return lines;
}

function startBatchTrailerVoice() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    const status = document.getElementById("batchVoiceStatus");
    const batchText = document.getElementById("batchText");

    if (!SpeechRecognition) {
        if (status) status.innerText = "Voice recognition is not supported in this browser. Try Chrome on Android.";
        return;
    }

    const recognition = new SpeechRecognition();
    recognition.lang = "en-GB";
    recognition.continuous = true;
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    if (status) status.innerText = "Listening... speak multiple trailer records, then pause.";

    recognition.onresult = function(event) {
        let text = "";
        for (let i = event.resultIndex; i < event.results.length; i++) {
            text += event.results[i][0].transcript + " ";
        }

        const lines = parseBulkTrailerSpeech(text);
        if (lines.length) {
            const existing = batchText.value.trim();
            batchText.value = existing ? existing + "\n" + lines.join("\n") : lines.join("\n");
            if (status) status.innerText = `${lines.length} trailer record(s) added from voice.`;
        } else {
            if (status) status.innerText = "Could not detect trailer records. Try: LG072 Door 12 loaded, DMC26020 Fence 18 empty.";
        }
    };

    recognition.onerror = function(event) {
        if (status) status.innerText = "Voice error: " + event.error;
    };

    recognition.onend = function() {
        if (status && status.innerText === "Listening... speak multiple trailer records, then pause.") {
            status.innerText = "Voice recording stopped.";
        }
    };

    recognition.start();
}
