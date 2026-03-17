/// <reference path="Xrm.d.ts" />
/**
 * Form-level notification banner for MPP import status on the adc_adccasetemplate form.
 *
 * Reads adc_importstatus (optionset) and adc_importmessage (string) fields
 * and shows/hides a setFormNotification banner accordingly.
 *
 * Register as onLoad event handler on the adc_adccasetemplate main form.
 * Function: ADC.CaseTemplateImportBanner.onLoad
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
ADC.CaseTemplateImportBanner = ADC.CaseTemplateImportBanner || {};

(function () {
    "use strict";

    var NOTIFICATION_ID = "template_import";
    var POLL_INTERVAL_MS = 15000; // 15 seconds
    var _intervalId = null;
    var _tickerId = null;
    var _startTime = null;

    // Status constants matching adc_importstatus optionset
    var STATUS = {
        QUEUED: 0,
        PROCESSING: 1,
        COMPLETED: 2,
        COMPLETED_WARNINGS: 3,
        FAILED: 4
    };

    /**
     * onLoad handler - shows banner and optionally starts polling.
     */
    ADC.CaseTemplateImportBanner.onLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Show initial banner
        updateBanner(formContext);

        var statusAttr = formContext.getAttribute("adc_importstatus");
        var status = statusAttr ? statusAttr.getValue() : null;

        if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
            // Import already in progress - poll immediately
            startPolling(formContext);
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
            var text = "Template project import in progress";
            if (msg) text += " - " + msg;

            if (!_startTime) {
                _startTime = new Date();
            }

            var elapsed = Math.floor((new Date() - _startTime) / 1000);
            var display = text + " (" + formatElapsed(elapsed) + ")";
            formContext.ui.setFormNotification(display, "INFO", NOTIFICATION_ID);
            startTicker(formContext);
        } else if (status === STATUS.FAILED) {
            stopTicker();
            var textErr = "Template project import failed";
            if (msg) textErr += " - " + msg;
            formContext.ui.setFormNotification(textErr, "ERROR", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED_WARNINGS) {
            stopTicker();
            var textWarn = "Template project import completed with warnings";
            if (msg) textWarn += " - " + msg;
            formContext.ui.setFormNotification(textWarn, "WARNING", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED) {
            stopTicker();
            var textOk = "Template project import completed";
            if (msg) textOk += " - " + msg;
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
                            // Keep polling if QUEUED or PROCESSING
                        }
                    },
                    function () {
                        // Refresh failed - stop polling to avoid noise
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
            var statusAttr = formContext.getAttribute("adc_importstatus");
            var msgAttr = formContext.getAttribute("adc_importmessage");
            var status = statusAttr ? statusAttr.getValue() : null;

            if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
                var msg = msgAttr ? (msgAttr.getValue() || "") : "";
                var text = "Template project import in progress";
                if (msg) text += " - " + msg;
                var elapsed = _startTime ? Math.floor((new Date() - _startTime) / 1000) : 0;
                var display = text + " (" + formatElapsed(elapsed) + ")";
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
        _startTime = null;
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
