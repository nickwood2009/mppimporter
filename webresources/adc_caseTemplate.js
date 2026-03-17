/// <reference path="Xrm.d.ts" />
/**
 * General form script for the adc_adccasetemplate form.
 *
 * - Shows/hides the "Plan Template" tab based on whether a template project
 *   is linked (adc_templateproject) AND the import is complete (adc_importstatus).
 * - Expands the embedded sub-form component to full height.
 *
 * Register on the adc_adccasetemplate main form:
 *   1. onLoad:  ADC.CaseTemplateForm.onLoad  (pass execution context)
 */
var ADC = ADC || {};
ADC.CaseTemplateForm = ADC.CaseTemplateForm || {};

(function () {
    "use strict";

    var TAB_NAME = "tab_plan";

    // Import status constants matching adc_importstatus optionset
    var STATUS = {
        QUEUED: 0,
        PROCESSING: 1,
        COMPLETED: 2,
        COMPLETED_WARNINGS: 3,
        FAILED: 4
    };

    ADC.CaseTemplateForm.onLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Evaluate tab visibility on load
        showHidePlanTab(formContext);

        // Re-evaluate when the template project lookup changes
        var projectAttr = formContext.getAttribute("adc_templateproject");
        if (projectAttr) {
            projectAttr.addOnChange(function () { showHidePlanTab(formContext); });
        }

        // Re-evaluate when the import status changes
        var statusAttr = formContext.getAttribute("adc_importstatus");
        if (statusAttr) {
            statusAttr.addOnChange(function () { showHidePlanTab(formContext); });
        }

        // Expand the embedded component when the tab is opened
        var tab = formContext.ui.tabs.get(TAB_NAME);
        if (tab) {
            tab.addTabStateChange(function () {
                expandTemplateProjectGrid();
            });
        }
    };

    /**
     * Shows the Plan Template tab only when:
     *   - adc_templateproject has a value, AND
     *   - adc_importstatus is Completed (2) or Completed with Warnings (3)
     * Hides it otherwise.
     */
    function showHidePlanTab(formContext) {
        var tab = formContext.ui.tabs.get(TAB_NAME);
        if (!tab) return;

        var projectAttr = formContext.getAttribute("adc_templateproject");
        var hasProject = projectAttr && projectAttr.getValue() != null
            && projectAttr.getValue().length > 0;

        var statusAttr = formContext.getAttribute("adc_importstatus");
        var status = statusAttr ? statusAttr.getValue() : null;

        var importDone = (status === STATUS.COMPLETED || status === STATUS.COMPLETED_WARNINGS);

        var visible = hasProject && importDone;
        tab.setVisible(visible);

        // If we just made the tab visible, expand the embedded component
        if (visible) {
            expandTemplateProjectGrid();
        }
    }

    /**
     * Polls the DOM for the embedded form component div and expands it
     * to full height once found. Uses setInterval to wait for the
     * component to render (it loads asynchronously after the tab opens).
     */
    function expandTemplateProjectGrid() {
        var pollId = setInterval(function () {
            var targetDiv = parent.window.document.querySelector(
                'div[id*="adc_templateproject"][id*="modelFormComponent"]');

            if (targetDiv) {
                targetDiv.style.height = "100%";
                clearInterval(pollId);
            }
        }, 500);
    }

})();
