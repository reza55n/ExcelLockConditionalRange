# ExcelLockConditionalRange
**Prevent unwanted changes in conditional formatting, during editing current sheet.**

This script keeps current worksheet's conditional formattings' ranges the same as specified first time. Also when pasting new rules into worksheet, it detects and removes duplicate rules.

**Tip:** This script checks duplicates only depending on their descriptions and conditions, not formattings (such as fill or border), neither ranges (as their ranges already became same).


## How to use
1. Open `ExcelLockConditionalRange.bas` file, copy script and paste it into intended worksheet's VBA code

2. (Optional) Change range and other settings in pasted script's first lines  
   (Range must be entered absolute, similar to the formula entered in Rules Manager window.)

3. Run `Conditional_ToggleLock()`


## Known issues
- Merging or unmerging cells won't fire `Worksheet_Change` event, therefore it's necessary to:
   - Make another change,
   - Switch to another worksheet and switch back,
   - Or run `Conditional_Refresh()` manually,

  to update rules.

- Similarly, editing conditional formatting rules or creating new ones won't fire `Worksheet_Change` event.

