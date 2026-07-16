/* clipgen Transcripts corrections satellite — transcripts-corrections.js
 *
 * The global find→replace "corrections" modal: list, add, and remove rules that
 * rewrite transcript text on the server. Loaded after transcripts.js; reads the
 * hub's shared state + helpers through window.ClipgenTranscripts (TS) and
 * publishes initCorrectionsModal / loadCorrections back so the hub's boot and the
 * inline-edit saveCorrections() flow can reach them. Plain utils.js globals
 * (qs/apiGet/apiPost/apiDelete/escapeHtml) are reached via the scope chain.
 */
(function () {
  "use strict";

  var TS = window.ClipgenTranscripts;
  var state = TS.state;
  var showToast = TS.showToast,
    loadTranscript = TS.loadTranscript;

  function closeCorrectionsModal() {
    var modal = qs("#correctionsModal");
    closeBlockingModal(modal);
    modal.classList.add("hidden");
  }

  function initCorrectionsModal() {
    qs("#correctionsBtn").addEventListener("click", function () {
      var modal = qs("#correctionsModal");
      modal.classList.remove("hidden");
      // Escape and backdrop click both dismiss (blocking-modal lifecycle also
      // keeps page hotkeys suppressed while the modal is up).
      openBlockingModal(modal, {
        onEscape: closeCorrectionsModal,
        onBackdropClick: closeCorrectionsModal,
      });
      loadCorrections();
    });

    qs("#closeCorrectionsBtn").addEventListener("click", closeCorrectionsModal);

    qs("#addCorrectionBtn").addEventListener("click", function () {
      addCorrection();
    });

    // Enter key in correction form
    qs("#correctionTo").addEventListener("keydown", function (e) {
      if (e.key === "Enter") {
        e.preventDefault();
        addCorrection();
      }
    });
  }

  function loadCorrections() {
    apiGet("api/corrections").then(function (data) {
      if (!data.ok) return;
      state.corrections = data.corrections;
      renderCorrections();
    });
  }

  function renderCorrections() {
    var container = qs("#correctionsList");
    if (state.corrections.length === 0) {
      container.innerHTML = '<div style="color:var(--color-text-dim);font-size:var(--text-sm);padding:var(--space-2) 0">No corrections yet</div>';
      return;
    }

    var html = "";
    state.corrections.forEach(function (c) {
      html += '<div class="correction-row">';
      html += '<span class="correction-from">' + escapeHtml(c.from) + '</span>';
      html += '<span class="correction-arrow">&rarr;</span>';
      html += '<span class="correction-to">' + escapeHtml(c.to) + '</span>';
      html += '<button class="correction-delete" data-id="' + escapeHtml(c.id) + '">Remove</button>';
      html += '</div>';
    });
    container.innerHTML = html;

    // Attach delete handlers
    var btns = container.querySelectorAll(".correction-delete");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", function () {
        deleteCorrection(this.getAttribute("data-id"));
      });
    }
  }

  function addCorrection() {
    var fromInput = qs("#correctionFrom");
    var toInput = qs("#correctionTo");
    var fromText = fromInput.value.trim();
    var toText = toInput.value.trim();
    if (!fromText || !toText) return;

    apiPost("api/corrections", { from: fromText, to: toText }).then(function (data) {
      if (data.ok) {
        fromInput.value = "";
        toInput.value = "";
        showToast("Correction added");
        loadCorrections();
        // Reload transcript to apply new correction
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  function deleteCorrection(id) {
    apiDelete("api/corrections/" + id).then(function (data) {
      if (data.ok) {
        showToast("Correction removed");
        loadCorrections();
        // Reload transcript to unapply correction
        if (state.selectedParticipant) loadTranscript(state.selectedParticipant);
      }
    });
  }

  // ---- Published back to the hub (boot + inline-edit saveCorrections call these) ----
  TS.initCorrectionsModal = initCorrectionsModal;
  TS.loadCorrections = loadCorrections;
})();
