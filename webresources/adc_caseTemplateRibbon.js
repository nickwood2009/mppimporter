/// <reference path="Xrm.d.ts" />
/**
 * Ribbon command functions for the adc_adccasetemplate form.
 *
 * Ribbon Workbench setup:
 *   - Add a custom button to the adc_adccasetemplate form command bar
 *   - Command action: ADC.CaseTemplateRibbon.importTemplate
 *   - CrmParameter: PrimaryControl
 *
 * IMPORTANT: Update IMPORT_TEMPLATE_WORKFLOW_ID with the actual workflow GUID
 * from your environment (Processes > Import Template > URL contains the id).
 */
var ADC = ADC || {};
ADC.CaseTemplateRibbon = ADC.CaseTemplateRibbon || {};

(function () {
    "use strict";

    // *** UPDATE THIS with your actual workflow ID ***
    var IMPORT_TEMPLATE_WORKFLOW_ID = "00000000-0000-0000-0000-000000000000";

    var NOTIFICATION_ID = "template_import";

    /**
     * Ribbon button action — triggers the ImportTemplate workflow for the
     * current case template record, then shows the banner and polls for updates.
     * @param {object} primaryControl - The form context (CrmParameter: PrimaryControl)
     */
    ADC.CaseTemplateRibbon.importTemplate = function (primaryControl) {
        var formContext = primaryControl;
        var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");

        // Confirm with user
        Xrm.Navigation.openConfirmDialog({
            title: "Import Template Project",
            text: "This will parse the MPP file and create a template project. Continue?"
        }).then(function (result) {
            if (!result.confirmed) return;

            // Show immediate feedback
            formContext.ui.setFormNotification(
                "Starting template project import...", "INFO", NOTIFICATION_ID);

            // Execute the workflow via Web API
            var requestUrl = Xrm.Utility.getGlobalContext().getClientUrl() +
                "/api/data/v9.2/workflows(" + IMPORT_TEMPLATE_WORKFLOW_ID + ")/Microsoft.Dynamics.CRM.ExecuteWorkflow";

            var requestBody = JSON.stringify({ EntityId: recordId });

            var req = new XMLHttpRequest();
            req.open("POST", requestUrl, true);
            req.setRequestHeader("Content-Type", "application/json");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");

            req.onreadystatechange = function () {
                if (req.readyState !== 4) return;

                if (req.status === 200 || req.status === 204) {
                    // Workflow started successfully — show banner and start polling
                    formContext.ui.setFormNotification(
                        "Template project import in progress — please wait...", "INFO", NOTIFICATION_ID);

                    // Start our own polling loop that refreshes form data and
                    // checks import status. We don't delegate to the banner's
                    // onChange because it clears the notification when status is
                    // still null (server hasn't processed yet).
                    startImportPolling(formContext);
                } else {
                    // Workflow failed to start
                    var errorMsg = "Failed to start import workflow.";
                    try {
                        var errorResponse = JSON.parse(req.responseText);
                        if (errorResponse.error && errorResponse.error.message) {
                            errorMsg += " " + errorResponse.error.message;
                        }
                    } catch (e) { /* ignore parse error */ }

                    formContext.ui.setFormNotification(errorMsg, "ERROR", NOTIFICATION_ID);

                    // Clear error after 10 seconds
                    setTimeout(function () {
                        formContext.ui.clearFormNotification(NOTIFICATION_ID);
                    }, 10000);
                }
            };

            req.send(requestBody);
        });
    };

    var _pollId = null;
    var POLL_INTERVAL_MS = 5000; // 5 seconds

    var STATUS = {
        QUEUED: 0,
        PROCESSING: 1,
        COMPLETED: 2,
        COMPLETED_WARNINGS: 3,
        FAILED: 4
    };

    /**
     * Polls the form data every 5 seconds after the workflow is triggered.
     * Keeps the banner alive until a real status appears, then hands off
     * to the import banner's updateBanner logic.
     */
    function startImportPolling(formContext) {
        if (_pollId) return; // already polling

        _pollId = setInterval(function () {
            try {
                formContext.data.refresh(false).then(function () {
                    var statusAttr = formContext.getAttribute("adc_importstatus");
                    var status = statusAttr ? statusAttr.getValue() : null;

                    if (status !== null && status !== undefined) {
                        // Server has set a real status — hand off to the banner JS
                        clearInterval(_pollId);
                        _pollId = null;

                        if (typeof ADC.CaseTemplateImportBanner !== "undefined" &&
                            typeof ADC.CaseTemplateImportBanner.onChange === "function") {
                            var fakeCtx = { getFormContext: function () { return formContext; } };
                            ADC.CaseTemplateImportBanner.onChange(fakeCtx);
                        }
                    }
                    // status is still null — keep polling, keep banner visible
                }, function () {
                    clearInterval(_pollId);
                    _pollId = null;
                });
            } catch (e) {
                clearInterval(_pollId);
                _pollId = null;
            }
        }, POLL_INTERVAL_MS);
    }

    /**
     * Enable rule — only enable the button when:
     *   - The record is saved (not a new unsaved record)
     *   - Import is not already in progress
     */
    ADC.CaseTemplateRibbon.enableImportButton = function (primaryControl) {
        var formContext = primaryControl;

        // Disable on create forms
        if (formContext.ui.getFormType() === 1) return false;

        // Disable if import is already running
        var statusAttr = formContext.getAttribute("adc_importstatus");
        if (statusAttr) {
            var status = statusAttr.getValue();
            if (status === 0 || status === 1) return false; // Queued or Processing
        }

        return true;
    };

})();
