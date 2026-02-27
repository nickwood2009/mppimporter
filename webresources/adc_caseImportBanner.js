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

        // Start polling if import is in progress
        var statusAttr = formContext.getAttribute("adc_importstatus");
        if (statusAttr) {
            var status = statusAttr.getValue();
            if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
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
            formContext.ui.setFormNotification(text, "INFO", NOTIFICATION_ID);
        } else if (status === STATUS.FAILED) {
            var textErr = "MPP import failed";
            if (msg) textErr += " — " + msg;
            formContext.ui.setFormNotification(textErr, "ERROR", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED_WARNINGS) {
            var textWarn = "MPP import completed with warnings";
            if (msg) textWarn += " — " + msg;
            formContext.ui.setFormNotification(textWarn, "WARNING", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED) {
            // Show success briefly then clear
            var textOk = "MPP import completed";
            if (msg) textOk += " — " + msg;
            formContext.ui.setFormNotification(textOk, "INFO", NOTIFICATION_ID);
            setTimeout(function () {
                formContext.ui.clearFormNotification(NOTIFICATION_ID);
            }, 10000);
            // Stop polling — import is done
            stopPolling();
        } else {
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
                Xrm.Page.data.refresh(false).then(
                    function () {
                        updateBanner(formContext);

                        // Stop polling if import is no longer in progress
                        var statusAttr = formContext.getAttribute("adc_importstatus");
                        if (statusAttr) {
                            var status = statusAttr.getValue();
                            if (status !== STATUS.QUEUED && status !== STATUS.PROCESSING) {
                                stopPolling();
                            }
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
     * Stops the polling interval.
     */
    function stopPolling() {
        if (_intervalId) {
            clearInterval(_intervalId);
            _intervalId = null;
        }
    }

})();
