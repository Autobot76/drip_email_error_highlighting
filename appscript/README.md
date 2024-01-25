# AppScript

---

This directory contains all the code in relation to error highlighting of the
Google Sheets.

It should be noted that this code will not run by itself. It is actually
embedded into the Google Sheets where it is tested. The code here should still
be considered the latest and most up-to-date version.

## Add Row/s Button (Templates)

---

### Known Issue

Clicking the "Add row" button will not cause the new row's fields to be
highlighted even though they are empty and are therefore invalid.

During testing this was discovered to only occur within the Templates Sheet.

Before jumping to any conclusions, it is recommended to give Google Sheets a
bit of time to process your request. It has been shown in testing that
sometimes Google Sheets is slow to update the highlighting. Ensure this is not
the case before proceeding below.

If at any point the user feels that it is not Google Sheets being slow they can
use the "Error Checker" button in the Google Sheets.

## Email Checking Edge case (Fixed)

---

asd@x.c is clearly an invalid email. The program will flag this as incorrect
but will not highlight this row.

This comes down to the fact that the regex used to identify a valid email has
not been accounted for.