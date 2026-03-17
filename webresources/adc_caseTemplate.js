/// <reference path="Xrm.d.ts" />
/**
 * Expands the embedded sub-form (related form component) on the adc_adccasetemplate form.
 *
 * Register on the adc_adccasetemplate main form:
 *   1. onLoad:  ADC.CaseTemplateForm.onLoad  (pass execution context)
 */
var ADC = ADC || {};
ADC.CaseTemplateForm = ADC.CaseTemplateForm || {};

(function () {
    "use strict";

    ADC.CaseTemplateForm.onLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();

        // Expand on initial load if the tab is already visible
        expandTemplateProjectGrid();

        // Also expand when the tab state changes (user clicks the tab)
        var tab = formContext.ui.tabs.get("tab_plan");
        if (tab) {
            tab.addTabStateChange(function () {
                ADC.CaseTemplateForm.showHideTemplateProjectGrid(formContext);
                expandTemplateProjectGrid();
            });
        }
    };

    /**
     * Show/hide logic for the embedded grid based on form state.
     * Extend as needed (e.g. hide grid on create forms).
     */
    ADC.CaseTemplateForm.showHideTemplateProjectGrid = function (formContext) {
        // Placeholder — add visibility logic here if needed
    };

    /**
     * Polls the DOM for the embedded form component div and expands it
     * to full height once found. Uses setInterval to wait for the
     * component to render (it loads asynchronously after the tab opens).
     */
    function expandTemplateProjectGrid() {
        setTimeout(function () {
            var pollId = setInterval(function () {
                // Look for the embedded model-driven form component div.
                // Update the selector below to match the actual div ID on your form.
                var targetDiv = parent.window.document.querySelector(
                    'div[id*="tab_sectionmk_projecttask_model|formComponent_0.modelFormComponent_0.modelFormComponent_"]'
                );

                if (targetDiv) {
                    targetDiv.style.height = "100%";
                    clearInterval(pollId);
                } else {
                    // Component not rendered yet — keep polling
                }
            }, 1000);

            // Safety: stop polling after 30 seconds to avoid infinite loop
            setTimeout(function () {
                clearInterval(pollId);
            }, 30000);
        }, 800);
    }

})();
