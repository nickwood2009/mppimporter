/// <reference path="Xrm.d.ts" />
/**
 * Form-level notification banner for MPP import status on the adc_case form.
 *
 * Reads adc_importstatus (optionset) and adc_importmessage (string) fields
 * and shows/hides a setFormNotification banner accordingly.
 *
 * Register as onLoad event handler on the adc_case main form.
 * Function: ADC.CaseImportBanner.onLoad
 * Pass execution context: Yes
 *
 * Status values:
 *   0 = Queued
 *   1 = Processing
 *   2 = Completed
 *   3 = Completed with Warnings
 *   4 = Failed
 */
var ADC = ADC || {};
ADC.CaseImportBanner = ADC.CaseImportBanner || {};

(function () {
    "use strict";

    var NOTIFICATION_ID = "mpp_import";
    var POLL_INTERVAL_MS = 15000; // 15 seconds
    var _intervalId = null;
    var _tickerId = null;
    var _lastMessage = "";
    var _elapsedSeconds = 0;

    // Status constants matching adc_importstatus optionset
    var STATUS = {
        QUEUED: 0,
        PROCESSING: 1,
        COMPLETED: 2,
        COMPLETED_WARNINGS: 3,
        FAILED: 4
    };

    /**
     * onLoad handler — shows banner and optionally starts polling.
     */
    ADC.CaseImportBanner.onLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Show initial banner
        updateBanner(formContext);

        var statusAttr = formContext.getAttribute("adc_importstatus");
        var status = statusAttr ? statusAttr.getValue() : null;

        if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
            // Import already in progress — poll immediately
            startPolling(formContext);
        } else if (status === null || status === undefined) {
            // Status not set yet — check if a case template is selected,
            // meaning the async plugin will kick off shortly
            var templateAttr = formContext.getAttribute("adc_adccasetemplateid");
            var hasTemplate = templateAttr && templateAttr.getValue() && templateAttr.getValue().length > 0;
            if (hasTemplate) {
                formContext.ui.setFormNotification(
                    "MPP import starting — please wait...", "INFO", NOTIFICATION_ID);
                startPolling(formContext);
            }
        }
    };

    /**
     * Reads import status fields and sets/clears the form notification.
     */
    function updateBanner(formContext) {
        var statusAttr = formContext.getAttribute("adc_importstatus");
        var msgAttr = formContext.getAttribute("adc_importmessage");

        if (!statusAttr) return;

        var status = statusAttr.getValue();
        var msg = msgAttr ? (msgAttr.getValue() || "") : "";

        if (status === null || status === undefined) {
            formContext.ui.clearFormNotification(NOTIFICATION_ID);
            return;
        }

        if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
            var text = "MPP import in progress";
            if (msg) text += " — " + msg;

            // Reset timer when server message changes
            if (msg !== _lastMessage) {
                _lastMessage = msg;
                _elapsedSeconds = 0;
            }

            var display = text + " (" + formatElapsed(_elapsedSeconds) + ")";
            formContext.ui.setFormNotification(display, "INFO", NOTIFICATION_ID);
            startTicker(formContext);
        } else if (status === STATUS.FAILED) {
            stopTicker();
            var textErr = "MPP import failed";
            if (msg) textErr += " — " + msg;
            formContext.ui.setFormNotification(textErr, "ERROR", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED_WARNINGS) {
            stopTicker();
            var textWarn = "MPP import completed with warnings";
            if (msg) textWarn += " — " + msg;
            formContext.ui.setFormNotification(textWarn, "WARNING", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED) {
            stopTicker();
            var textOk = "MPP import completed";
            if (msg) textOk += " — " + msg;
            formContext.ui.setFormNotification(textOk, "INFO", NOTIFICATION_ID);
            setTimeout(function () {
                formContext.ui.clearFormNotification(NOTIFICATION_ID);
            }, 10000);
            stopPolling();
        } else {
            stopTicker();
            formContext.ui.clearFormNotification(NOTIFICATION_ID);
            stopPolling();
        }
    }

    /**
     * Starts a polling interval that refreshes data and updates the banner.
     */
    function startPolling(formContext) {
        if (_intervalId) return; // already polling

        _intervalId = setInterval(function () {
            try {
                // Refresh the form data to pick up server-side changes
                formContext.data.refresh(false).then(
                    function () {
                        updateBanner(formContext);

                        // Stop polling only when import reaches a terminal state
                        var statusAttr = formContext.getAttribute("adc_importstatus");
                        if (statusAttr) {
                            var status = statusAttr.getValue();
                            if (status === STATUS.COMPLETED ||
                                status === STATUS.COMPLETED_WARNINGS ||
                                status === STATUS.FAILED) {
                                stopPolling();
                            }
                            // Keep polling if null (waiting for async plugin),
                            // QUEUED, or PROCESSING
                        }
                    },
                    function () {
                        // Refresh failed — stop polling to avoid noise
                        stopPolling();
                    }
                );
            } catch (e) {
                stopPolling();
            }
        }, POLL_INTERVAL_MS);
    }

    /**
     * Starts a 1-second ticker that increments the elapsed counter
     * and updates the banner text so it looks live.
     */
    function startTicker(formContext) {
        if (_tickerId) return; // already ticking

        _tickerId = setInterval(function () {
            _elapsedSeconds++;
            var statusAttr = formContext.getAttribute("adc_importstatus");
            var msgAttr = formContext.getAttribute("adc_importmessage");
            var status = statusAttr ? statusAttr.getValue() : null;

            if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
                var msg = msgAttr ? (msgAttr.getValue() || "") : "";
                var text = "MPP import in progress";
                if (msg) text += " — " + msg;
                var display = text + " (" + formatElapsed(_elapsedSeconds) + ")";
                formContext.ui.setFormNotification(display, "INFO", NOTIFICATION_ID);
            }
        }, 1000);
    }

    /**
     * Stops the 1-second ticker.
     */
    function stopTicker() {
        if (_tickerId) {
            clearInterval(_tickerId);
            _tickerId = null;
        }
        _elapsedSeconds = 0;
        _lastMessage = "";
    }

    /**
     * Formats seconds into a human-readable elapsed string (e.g. "1m 23s").
     */
    function formatElapsed(totalSeconds) {
        var mins = Math.floor(totalSeconds / 60);
        var secs = totalSeconds % 60;
        if (mins > 0) {
            return mins + "m " + secs + "s";
        }
        return secs + "s";
    }

    /**
     * Stops the polling interval.
     */
    function stopPolling() {
        if (_intervalId) {
            clearInterval(_intervalId);
            _intervalId = null;
        }
        stopTicker();
    }

})();
