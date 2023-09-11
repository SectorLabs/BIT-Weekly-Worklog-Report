# GAS-BIT-Weekly-Report

This GitHub repository contains code for automating the generation of a weekly report for a team. The script is written in Google Apps Script and is designed to work with Google Sheets, Google Calendar, and Gmail.

**Features:**

BITPK_weekclickuptime(): This function retrieves data from ClickUp using API keys and populates it in a Google Sheet named 'WeeklyTeamClickUp'. It calculates time spent on tasks within a specified week, filtering data between Monday and Friday.

BIT_WeeklySDPRequests(): This function fetches service requests from an external source using API calls and populates the data in a Google Sheet named 'Weekly SDP'. It retrieves requests updated during the specified week.

BIT_WeeklyWorklog(): This function complements the weekly report by fetching worklog data related to service requests and populates it in a Google Sheet named 'WeeklyTeamWorklog'. It retrieves worklog entries within the specified week.

sendWeeklyReport(): This function compiles a weekly report for team members by combining data from the 'WeeklyTeamClickUp', 'Weekly SDP', and Google Calendar. It calculates total time spent on tasks and meetings, then sends the report as a PDF via email.

getWeekOfMonth(): A helper function to determine the week number of the month based on a given date.

**Usage:**

To use this code:

Create a Google Sheet with the following sheet names: 'WeeklyTeamClickUp', 'Weekly SDP', and 'WeeklyTeamWorklog'.

Set up external API keys as required for ClickUp and Zoho OAuth token for SDP API calls.

Copy and paste the code into Google Apps Script.

Run the sendWeeklyReport() function to generate and send the weekly report.

Schedule the script to run automatically at your desired frequency using Google Apps Script triggers.

Note: Ensure that you have the necessary permissions and API access for the external services (ClickUp, SDP, and Google Calendar) to retrieve data.

