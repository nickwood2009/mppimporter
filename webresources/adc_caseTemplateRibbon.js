/// <reference path="Xrm.d.ts" />
/**
 * Ribbon command functions for the adc_adccasetemplate form.
 *
 * Ribbon Workbench setup:
 *   - Add a custom button to the adc_adccasetemplate form command bar
 *   - Command action: ADC.CaseTemplateRibbon.importTemplate
 *   - CrmParameter: PrimaryControl
 *   - Display Rule: ADC.CaseTemplateRibbon.isAllowedRole (PO Business Admin or System Administrator)
 *
 * Calls the "adc_ImportMppTemplate" Custom Action (bound to adc_adccasetemplate).
 * The old on-demand workflow can be deactivated/deleted.
 */
var ADC = ADC || {};
ADC.CaseTemplateRibbon = ADC.CaseTemplateRibbon || {};

(function () {
    "use strict";

    // Custom Action unique name (bound to adc_adccasetemplate entity)
    var IMPORT_ACTION_NAME = "adc_ImportMppTemplate";

    var NOTIFICATION_ID = "template_import";

    /**
     * Ribbon button action — triggers the ImportTemplate workflow for the
     * current case template record, then shows the banner and polls for updates.
     * @param {object} primaryControl - The form context (CrmParameter: PrimaryControl)
     */
    ADC.CaseTemplateRibbon.importTemplate = function (primaryControl) {
        var formContext = primaryControl;
        var recordId = formContext.data.entity.getId().replace(/[{}]/g, "");

        // Guard — block if user doesn't have an allowed role
        userHasAllowedRole().then(function (hasRole) {
            if (!hasRole) {
                Xrm.Navigation.openAlertDialog({
                    title: "Access Denied",
                    text: "You must have the 'PO Business Admin' or 'System Administrator' security role to run this import."
                });
                return;
            }
            _runImport(formContext, recordId);
        });
    };

    function _runImport(formContext, recordId) {
        // Confirm with user
        Xrm.Navigation.openConfirmDialog({
            title: "Import Template Project",
            text: "This will parse the MPP file and create a template project. Continue?"
        }).then(function (result) {
            if (!result.confirmed) return;

            // Show immediate feedback
            formContext.ui.setFormNotification(
                "Starting template project import...", "INFO", NOTIFICATION_ID);

            // Execute the bound Custom Action via Web API
            var requestUrl = Xrm.Utility.getGlobalContext().getClientUrl() +
                "/api/data/v9.2/adc_adccasetemplates(" + recordId + ")/Microsoft.Dynamics.CRM." + IMPORT_ACTION_NAME;

            var requestBody = JSON.stringify({});

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
    }

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

    // ──────────────────────────────────────────────────────────
    // Security role check — "PO Business Admin" or "System Administrator"
    // ──────────────────────────────────────────────────────────

    // Allowed roles — checked by GUID first, falls back to name if GUID is placeholder.
    // *** UPDATE GUIDs with the actual values from your environment ***
    // GUIDs are preserved across environments when exported/imported via solutions.
    var ALLOWED_ROLES = [
        { id: "00000000-0000-0000-0000-000000000000", name: "PO Business Admin" },
        { id: "00000000-0000-0000-0000-000000000000", name: "System Administrator" }
    ];
    var PLACEHOLDER_GUID = "00000000-0000-0000-0000-000000000000";
    var _hasRoleCache = null; // cache result for the session

    /**
     * Checks whether the current user has any of the allowed security roles.
     * Tries matching by role GUID first; if the GUID is still a placeholder,
     * falls back to matching by name.
     * Result is cached so subsequent calls don't re-query.
     * @returns {Promise<boolean>}
     */
    function userHasAllowedRole() {
        if (_hasRoleCache !== null) {
            return Promise.resolve(_hasRoleCache);
        }

        var userId = Xrm.Utility.getGlobalContext().userSettings.userId
            .replace(/[{}]/g, "");
        var clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

        // Build OData filter: match by ID or name for each allowed role
        var filters = [];
        for (var i = 0; i < ALLOWED_ROLES.length; i++) {
            var role = ALLOWED_ROLES[i];
            if (role.id && role.id !== PLACEHOLDER_GUID) {
                filters.push("roleid eq " + role.id);
            } else {
                filters.push("name eq '" + encodeURIComponent(role.name) + "'");
            }
        }

        var url = clientUrl +
            "/api/data/v9.2/systemusers(" + userId + ")/systemuserroles_association" +
            "?$select=roleid,name&$filter=" + filters.join(" or ");

        return new Promise(function (resolve) {
            var req = new XMLHttpRequest();
            req.open("GET", url, true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");

            req.onreadystatechange = function () {
                if (req.readyState !== 4) return;
                if (req.status === 200) {
                    try {
                        var data = JSON.parse(req.responseText);
                        _hasRoleCache = data.value && data.value.length > 0;
                    } catch (e) {
                        _hasRoleCache = false;
                    }
                } else {
                    _hasRoleCache = false;
                }
                resolve(_hasRoleCache);
            };
            req.send();
        });
    }

    /**
     * Display rule for Ribbon Workbench — controls button VISIBILITY.
     * Returns true if the user has "PO Business Admin" or "System Administrator".
     *
     * Ribbon Workbench setup:
     *   - Add a Display Rule of type "Custom Rule"
     *   - Library: adc_caseTemplateRibbon
     *   - FunctionName: ADC.CaseTemplateRibbon.isAllowedRole
     *   - CrmParameter: PrimaryControl
     */
    ADC.CaseTemplateRibbon.isAllowedRole = function () {
        return userHasAllowedRole().then(function (hasRole) {
            return hasRole;
        });
    };

})();
