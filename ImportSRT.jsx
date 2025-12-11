// ImportSRT.jsx
// After Effects ExtendScript (compatible with AE 2025)
// Imports an SRT and creates editable text layers.
// You can control how many words per subtitle layer by changing WORDS_PER_LAYER.
// Save as ImportSRT.jsx and run from File > Scripts > Run Script File...

(function importSRT() {
    // --------- User settings (edit these) ----------
    var WORDS_PER_LAYER = 3;      // 1 = one word per layer, 2 = two words per layer, etc.
    var FONT_NAME = "";          // e.g. "Arial-BoldMT" or "" to keep default font
    // ------------------------------------------------

    var undoStarted = false;
    try {
        if (app.project === null) {
            app.newProject();
        }

        var comp = app.project.activeItem;
        if (!(comp instanceof CompItem)) {
            alert("Please open or select a composition before running this script.");
            return;
        }

        // Ask user to pick SRT
        var srtFile = File.openDialog("Select an SRT file", "*.srt");
        if (srtFile === null) {
            return; // cancelled
        }

        // Ensure UTF-8 reading
        try {
            srtFile.encoding = "UTF-8";
        } catch (eEnc) {
            // some ExtendScript environments may ignore; continue anyway
        }
        if (!srtFile.open("r")) {
            alert("Unable to open file: " + srtFile.fsName);
            return;
        }
        var srtText = srtFile.read();
        srtFile.close();

        // Normalize line endings to \n
        srtText = String(srtText).replace(/\r\n/g, "\n").replace(/\r/g, "\n");

        var cues = parseSRT(srtText);
        if (!cues || cues.length === 0) {
            alert("No subtitle cues found in the SRT file.");
            return;
        }

        app.beginUndoGroup("Import SRT Subtitles");
        undoStarted = true;

        // Extend comp duration if needed
        var lastOut = cues[cues.length - 1]["out"];
        if (lastOut > comp.duration) {
            comp.duration = lastOut + 1;
        }

        // Create text layers
        for (var i = 0; i < cues.length; i++) {
            var cue = cues[i];
            var layer = comp.layers.addText(" ");
            layer.name = "Subtitle " + (i + 1);
            layer.inPoint = cue["in"];
            layer.outPoint = cue["out"];

            // Position at center of comp (X center, Y center)
            try {
                layer.property("Position").setValue([comp.width / 2, comp.height / 2]);
            } catch (ePos) {
                // ignore if property not found
            }

            // Fill Source Text
            var textProp = layer.property("Source Text");
            if (textProp) {
                var textDoc = textProp.value;
                // cue.text is already safe string
                textDoc.text = cue["text"];
                // Optional font
                if (FONT_NAME && FONT_NAME !== "") {
                    try { textDoc.font = FONT_NAME; } catch (eF) { /* ignore invalid font */ }
                }
                textDoc.fontSize = Math.max(20, Math.round(comp.height / 15));
                textDoc.leading = textDoc.fontSize * 1.05;
                try { textDoc.justification = ParagraphJustification.CENTER_JUSTIFY; } catch (eJ) { /* ignore */ }
                textProp.setValue(textDoc);
            }
        }

        app.endUndoGroup();
        undoStarted = false;
        alert("Imported " + cues.length + " subtitle layers.");

    } catch (err) {
        if (undoStarted) {
            try { app.endUndoGroup(); } catch (e) { /* ignore */ }
        }
        alert("Error: " + (err && err.toString ? err.toString() : String(err)));
    }

    // ---------------- Helper functions ----------------

    function parseSRT(text) {
        // Split by blank lines (two or more newlines)
        var items = String(text).split(/\n{2,}/);
        var out = [];
        for (var i = 0; i < items.length; i++) {
            var block = items[i];
            if (typeof block !== "string") {
                if (block === null || block === undefined) {
                    continue;
                } else {
                    block = String(block);
                }
            }
            // manual trim (ExtendScript-safe)
            block = block.replace(/^\s+|\s+$/g, "");
            if (block === "") continue;

            var lines = block.split("\n");

            // Find timecode line (flexible: allow it on first up to third line)
            var timeLineIndex = -1;
            var checkMax = Math.min(lines.length, 3);
            for (var k = 0; k < checkMax; k++) {
                if (lines[k].match(/\d{1,2}:\d{2}:\d{2}[,\.]\d{1,3}\s*-->\s*\d{1,2}:\d{2}:\d{2}[,\.]\d{1,3}/)) {
                    timeLineIndex = k;
                    break;
                }
            }
            if (timeLineIndex === -1) continue;

            var timeLine = lines[timeLineIndex];
            var m = timeLine.match(/(\d{1,2}:\d{2}:\d{2})[:,\.](\d{1,3})\s*-->\s*(\d{1,2}:\d{2}:\d{2})[:,\.](\d{1,3})/);
            if (!m) continue;

            var inSec = toSeconds(m[1], m[2]);
            var outSec = toSeconds(m[3], m[4]);

            // Join any text lines after the timecode into one string
            var textLines = lines.slice(timeLineIndex + 1);
            var cueText = textLines.join(" ");
            cueText = stripTags(cueText);
            // manual trim again
            cueText = cueText.replace(/^\s+|\s+$/g, "");

            if (cueText === "") continue;

            // Now split into groups of WORDS_PER_LAYER
            var words = cueText.split(/\s+/);
            // remove any empty items
            var cleanedWords = [];
            for (var wi = 0; wi < words.length; wi++) {
                if (typeof words[wi] !== "string") continue;
                var w = words[wi].replace(/^\s+|\s+$/g, "");
                if (w === "") continue;
                cleanedWords.push(w);
            }
            if (cleanedWords.length === 0) continue;

            var groupSize = (typeof WORDS_PER_LAYER === "number" && WORDS_PER_LAYER > 0) ? Math.floor(WORDS_PER_LAYER) : 1;
            if (groupSize < 1) groupSize = 1;

            var chunkCount = Math.ceil(cleanedWords.length / groupSize);
            var totalDuration = outSec - inSec;
            // avoid division by zero
            var chunkDuration = (chunkCount > 0) ? (totalDuration / chunkCount) : totalDuration;

            for (var c = 0; c < chunkCount; c++) {
                var startIndex = c * groupSize;
                var endIndex = Math.min(startIndex + groupSize, cleanedWords.length);
                var chunkWords = [];
                for (var ci = startIndex; ci < endIndex; ci++) {
                    chunkWords.push(cleanedWords[ci]);
                }
                var chunkText = chunkWords.join(" ");
                var win = inSec + c * chunkDuration;
                var wout = inSec + (c + 1) * chunkDuration;
                // push as an object with quoted keys (ExtendScript-safe)
                out.push({ "in": win, "out": wout, "text": chunkText });
            }
        }
        return out;
    }

    function toSeconds(hms, ms) {
        var parts = String(hms).split(":");
        var hh = parseInt(parts[0], 10) || 0;
        var mm = parseInt(parts[1], 10) || 0;
        var ss = parseInt(parts[2], 10) || 0;
        var msecs = parseInt((String(ms) + "000").substr(0, 3), 10) || 0;
        return hh * 3600 + mm * 60 + ss + msecs / 1000;
    }

    function stripTags(s) {
        if (s === null || s === undefined) return "";
        return String(s).replace(/<\/?[^>]+(>|$)/g, "");
    }

})();
