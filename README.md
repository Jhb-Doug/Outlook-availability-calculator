# Outlook-availability-calculator
VBA form allowing user to select a portion of the calendar and generate text indicating their availability within that period.
This calculator was inspired by many emails asking for my availability and having to manually check the calendar and type it out.

To use this calculator, open the VBA editor from Outlook and import the form and module.
Next go to Outlook File-->options-->customize ribbon (Suggest doing this under the Calendar category)
Add a new group to the list on the right side of the screen
Select macros from the left dropdown combo box, select the OpenAvailabilityCalculator macro and add a new command to the group you created.
Name it appropriately.

Once the button is added, select the portion of the calendar you want to generate availability for and click the button you created.
This should load the default dates and times in the calculator for the period which you selected. Edit any parameters you want to and
click the calculate button. Copy the text and paste it where you need it.

