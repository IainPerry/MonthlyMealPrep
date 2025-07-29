# MonthlyMealPrep
A google sheet set of scripts to automatically make meals for a month and generate a shopping list.

The xlsx holds the structure of google sheet you'll want to use/recreate.
The boxes "Generate" need assigning to a script. Add all the scripts to your Google Scripts project, and assign to the boxes to run. e.g. assign "generateAugust" to the box filling in the table for August.

Cells are specific so moving them around may break the script, but it should be fairly easy to tweak in the scripts to fit your preference.


# Areas for development
1. Automatically finding the first 'Monday' of the week as part of the shopping list generation.
2. Accounting for packaging that may want using up over a couple of days.
3. Dynamic generation of a print sheet for weekly menu and shopping list.
