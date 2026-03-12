# Duration Variance Fields — msdyn_projecttask

Add these 5 custom fields to the `msdyn_projecttask` entity.

## Fields

| Schema Name | Display Name | Type | Details |
|---|---|---|---|
| `adc_sourcedurationdays` | Source Duration (Days) | Whole Number | Min: 0, Max: 2147483647 |
| `adc_sourcedurationhours` | Source Duration (Hours) | Whole Number | Min: 0, Max: 2147483647 |
| `adc_durationvariancedays` | Duration Variance (Days) | Decimal | Precision: 2, Min: -100000, Max: 100000 |
| `adc_durationvariancereason` | Duration Variance Reason | Single Line of Text | Max Length: 200 |
| `adc_issourcemilestone` | Is Source Milestone | Yes/No | Default: No |

## Quick Reference (copy-paste)

```
Schema: adc_sourcedurationdays
Display: Source Duration (Days)
Type: Whole Number

Schema: adc_sourcedurationhours
Display: Source Duration (Hours)
Type: Whole Number

Schema: adc_durationvariancedays
Display: Duration Variance (Days)
Type: Decimal (precision 2)

Schema: adc_durationvariancereason
Display: Duration Variance Reason
Type: Single Line of Text (200)

Schema: adc_issourcemilestone
Display: Is Source Milestone
Type: Yes/No
```

## What they do

- **Source Duration (Days/Hours)** — Original MPP template duration, set during import
- **Duration Variance (Days)** — `msdyn_duration - adc_sourcedurationdays` (positive = PSS added time)
- **Duration Variance Reason** — Auto-populated explanation (e.g. "adjusted for non-working days", "Milestone - no duration", "Summary task - PSS auto-calculates from children")
- **Is Source Milestone** — True if MPP source had 0-duration (milestone task)
