Data validation can help control what a user can enter into a cell. You can use data validation to make sure a value is a number, a date, or to present a dropdown menu with predefined choices to a user. This guide provides an overview of the data validation feature, with many examples.Data validation is a feature in Excel used to control what a user can enter into a cell. For example, you could use data validation to make sure a value is a number between 1 and 6, make sure a date occurs in the next 30 days, or make sure a text entry is less than 25 characters.Important limitation
It is important to understand that data validation can be easily defeated. If a user copies data from a cell without validation to a cell with data validation, the validation is destroyed (or replaced). Data validation is a good way to let users know what is allowed or expected, but it is not a foolproof way to guarantee input.Data validation options
When a data validation rule is created, there are eight options available to validate user input:

Any Value - no validation is performed. Note: if data validation was previously applied with a set Input Message, the message will still display when the cell is selected, even when Any Value is selected.

Whole Number - only whole numbers are allowed. Once the whole number option is selected, other options become available to further limit input. For example, you can require a whole number between 1 and 10.

Decimal - works like the whole number option, but allows decimal values. For example, with the Decimal option configured to allow values between 0 and 3, values like .5, 2.5, and 3.1 are all allowed.

List - only values from a predefined list are allowed. The values are presented to the user as a dropdown menu control. Allowed values can be hardcoded directly into the Settings tab, or specified as a range on the worksheet.

Date - only dates are allowed. For example, you can require a date between January 1, 2018 and December 31 2021, or a date after June 1, 2018.

Time - only times are allowed. For example, you can require a time between 9:00 AM and 5:00 PM, or only allow times after 12:00 PM.

Text length - validates input based on number of characters or  digits. For example, you could require code that contains 5 digits.Custom - validates user input using a custom formula. In other words, you can write your own formula to validate input. Custom formulas greatly extend the options for data validation. For example, you could use a formula to ensure a value is uppercase, a value contains "xyz", or a date is a weekday in the next 45 days.

The settings tab also includes two checkboxes:

Ignore blank - tells Excel to not validate cells that contain no value. In practice, this setting seems to affect only the command "circle invalid data". When enabled, blank cells are not circled even if they fail validation.

Apply these changes to other cells with the same settings - this setting will update validation applied to other cells when it matches the (original) validation of the cell(s) being edited.

Note: You can also manually select all cells with data validation applied using Go To + Special, as explained below.

Simple drop down menu
You can provide a dropdown menu of options by hardcoding values into the settings box, or selecting a range on the worksheet. For example, to restrict entries to the actions "BUY", "HOLD", or "SELL"The possibilities for data validation custom formulas are virtually unlimited. Here are a few examples to give you some inspiration:

To allow only 5 character values that begin with "z" you could use:

=AND(LEFT(A1)="z",LEN(A1)=5)
This formula returns TRUE only when a code is 5 digits long and starts with "z". The two circled values return FALSE with this formula. 

To allow only a date within 30 days of today:

=AND(A1>TODAY(),A1<=(TODAY()+30))
To allow only unique values:

=COUNTIF(range,A1)<2
To allow only an email address

=ISUMBER(FIND("@",A1)
