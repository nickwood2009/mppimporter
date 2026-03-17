/// <reference path="Xrm.d.ts" />
/**
 * Form-level notification banner for MPP import status on the adc_case form.
 *
 * Reads adc_importstatus (optionset) and adc_importmessage (string) fields
 * and shows/hides a setFormNotification banner accordingly.
 *
 * Register on the adc_case main form:
 *   1. onLoad:   ADC.CaseImportBanner.onLoad   (pass execution context)
 *   2. onSave:    ADC.CaseImportBanner.onSave   (pass execution context)
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
     * onLoad handler — shows banner and optionally starts polling.
     */
    ADC.CaseImportBanner.onLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Show initial banner
        updateBanner(formContext);

        var statusAttr = formContext.getAttribute("adc_importstatus");
        var status = statusAttr ? statusAttr.getValue() : null;

        if (status === STATUS.QUEUED || status === STATUS.PROCESSING) {
            // Import/clone already in progress — poll immediately
            startPolling(formContext);
        } else if (status === null || status === undefined) {
            // Status not set yet — check if a case template is selected,
            // meaning the async workflow will kick off shortly
            var templateAttr = formContext.getAttribute("adc_adccasetemplateid");
            var hasTemplate = templateAttr && templateAttr.getValue() && templateAttr.getValue().length > 0;
            if (hasTemplate) {
                formContext.ui.setFormNotification(
                    "Project setup starting — please wait...", "INFO", NOTIFICATION_ID);
                startPolling(formContext);
            }
        }
    };

    /**
     * onSave handler — after the user saves the form, if a case template is
     * selected the async workflow will fire server-side. Show the banner and
     * start polling so the user sees progress without a manual refresh.
     */
    ADC.CaseImportBanner.onSave = function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Only kick off polling if a template is set and status hasn't reached a terminal state
        var templateAttr = formContext.getAttribute("adc_adccasetemplateid");
        var hasTemplate = templateAttr && templateAttr.getValue() && templateAttr.getValue().length > 0;
        if (!hasTemplate) return;

        var statusAttr = formContext.getAttribute("adc_importstatus");
        var status = statusAttr ? statusAttr.getValue() : null;

        // If already completed/failed, don't restart polling
        if (status === STATUS.COMPLETED || status === STATUS.COMPLETED_WARNINGS || status === STATUS.FAILED) return;

        // Delay banner display until after the post-save form re-render,
        // otherwise Dynamics clears the notification during its re-render cycle.
        setTimeout(function () {
            _startTime = new Date();
            formContext.ui.setFormNotification(
                "Project setup starting — please wait...", "INFO", NOTIFICATION_ID);
            startTicker(formContext);
            startPolling(formContext);
        }, 2000);
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
            var text = "Project setup in progress";
            if (msg) text += " — " + msg;

            if (!_startTime) {
                _startTime = new Date();
            }

            var elapsed = Math.floor((new Date() - _startTime) / 1000);
            var display = text + " (" + formatElapsed(elapsed) + ")";
            formContext.ui.setFormNotification(display, "INFO", NOTIFICATION_ID);
            startTicker(formContext);
        } else if (status === STATUS.FAILED) {
            stopTicker();
            var textErr = "Project setup failed";
            if (msg) textErr += " — " + msg;
            formContext.ui.setFormNotification(textErr, "ERROR", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED_WARNINGS) {
            stopTicker();
            var textWarn = "Project setup completed with warnings";
            if (msg) textWarn += " — " + msg;
            formContext.ui.setFormNotification(textWarn, "WARNING", NOTIFICATION_ID);
        } else if (status === STATUS.COMPLETED) {
            stopTicker();
            var textOk = "Project setup completed";
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
                            } else if (status === STATUS.PROCESSING) {
                                // Server status stuck at Processing — check if the
                                // linked project copy has actually finished
                                checkProjectCopyStatus(formContext);
                            }
                            // Keep polling if null (waiting for async plugin) or QUEUED
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
     * Checks if the linked project's copy operation has completed by reading
     * the project's statuscode via Web API. If Active (1), marks the case as
     * Completed so the banner updates without waiting for a server-side process.
     */
    function checkProjectCopyStatus(formContext) {
        var projectAttr = formContext.getAttribute("adc_projectid");
        if (!projectAttr) return;
        var projectVal = projectAttr.getValue();
        if (!projectVal || projectVal.length === 0) return;

        var projectId = projectVal[0].id.replace(/[{}]/g, "");

        Xrm.WebApi.retrieveRecord("msdyn_project", projectId, "?$select=statuscode").then(
            function (result) {
                var statusCode = result.statuscode;
                // statuscode 1 = Active (copy done)
                if (statusCode === 1) {
                    // Update case import status to Completed
                    var caseId = formContext.data.entity.getId().replace(/[{}]/g, "");
                    var updateData = {
                        adc_importstatus: STATUS.COMPLETED,
                        adc_importmessage: "Project cloned successfully."
                    };
                    Xrm.WebApi.updateRecord("adc_case", caseId, updateData).then(
                        function () {
                            // Refresh form to pick up the change
                            formContext.data.refresh(false).then(function () {
                                updateBanner(formContext);
                                stopPolling();
                            });
                        },
                        function () { /* non-fatal */ }
                    );
                }
                // If statuscode == 192350000 (copy failed), mark as failed
                else if (statusCode === 192350000) {
                    var caseId2 = formContext.data.entity.getId().replace(/[{}]/g, "");
                    var failData = {
                        adc_importstatus: STATUS.FAILED,
                        adc_importmessage: "Project copy failed. Check PSS Error Logs."
                    };
                    Xrm.WebApi.updateRecord("adc_case", caseId2, failData).then(
                        function () {
                            formContext.data.refresh(false).then(function () {
                                updateBanner(formContext);
                                stopPolling();
                            });
                        },
                        function () { /* non-fatal */ }
                    );
                }
                // Otherwise still copying — keep polling
            },
            function () { /* Web API call failed — keep polling */ }
        );
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

            if (status === STATUS.QUEUED || status === STATUS.PROCESSING || status === null) {
                var msg = msgAttr ? (msgAttr.getValue() || "") : "";
                var text = (status === null) ? "Project setup starting" : "Project setup in progress";
                if (msg) text += " — " + msg;
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
