# PAM

Privileged Access Management Review Report Generator
Version 3

Description:
This script compares the deltas between the BEFORE and AFTER Privileged Security Group reports.

It also reports user-based roles assigned to non-ISC workers for review.

Requirements:
On the same working folder as this notebook, store the before and after Privileged Security Group excel files. The file name format for these reports should be: PSGyyyymmdd.xlsx

The recommended working folder to use is: C:\Users\your-user-name\Documents\PAM

Then, in Parameters below, update the fn_bfore and fn_after with the appropriate file names, without the '.xlsx'

Parameters:  Before and After files. Example Below.
<ul>
fn_bfore = 'PSG20220713'
fn_after = 'PSG20220822'
</ul>
