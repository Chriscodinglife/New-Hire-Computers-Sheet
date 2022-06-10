/*

For reference to Google Apps Scripts Documentation visit here: https://developers.google.com/apps-script/overview
Chris O
2022

// :: About::

This is a google/javascript that is intended to use in conjunction with keeping track of new hires and their fedex tracking numbers.
It's crucial function is main() that takes in the parameter e as the event argument.
When an event happens the following is handle by this Apps Script:
1. If a user enters a a work email under the Work email column, the Full name and Username will be automatically set
2. If the user sets the computer Shipped status to Yes, an email will be automatically sent out to the new hire with their tracking number
3. If a user sets the Send Reminder column to Yes, an email will be sent out to remind the new that they need to setup their comptuer before their start date.

*/

// VARIABLES ////////////////////////////////
/////////////////////////////////////////////
// Hard Code the position of the columns for sanity
// Please note that the index has a starting position of 1 and not 0

// EDIT THESE VALUES AS NEEDED
var state_date_column = 1;
var full_name_column = 2;
var personal_email_column = 3;
var user_name_column = 4;
var work_email_column = 5;
var office_location_column = 6;
var shipping_number_column = 7;
var computer_shipped_status_column = 8;
var setup_email_sent_status_column = 9;
var send_reminder_column = 10;
var reminder_email_sent_status_column = 11;

var email_sent = "EMAIL SENT";
var reminder_sent = "REMINDER SENT";

const bcc_email = "<your_bcc_email>";
const company_domain = "<your_@_company_domain_here>";
const it_team_name = "<your_team_name>";
const company_name = "<your_company_name>";
const current_tab_name = "<your_google_sheet_tab_name>"

const tracking_url = "https://www.fedex.com/fedextrack/?trknbr=";

var ui = SpreadsheetApp.getUi(); // Instantiate the UI prompt and name it "ui"
/////////////////////////////////////////////
/////////////////////////////////////////////


// FUNCTIONS ////////////////////////////////
function is_current_spreadsheet_this_sheet(sheet_name) {

    // Return true if the current spreadsheet being edited is the given sheet name
    var current_sheet = SpreadsheetApp.getActiveSheet();
  
    if (current_sheet.getSheetName() === sheet_name) {
      return true;
    }
  
  }
  
  function return_current_sheet() {
  
    // Return the current spreadsheet
    var current_sheet = SpreadsheetApp.getActiveSheet();
    return current_sheet;
  
  }
  
  function get_edited_row(this_event, sheet) {
  
    // Return a list and index value of the row of the edited cell from the current_sheet
    var row_index = this_event.range.getRow();
    var row_list = sheet.getRange(row_index, 1, 1, sheet.getLastColumn()).getValues()[0];
  
    return [row_index, row_list];
  
  }
  
  function get_username_info(email) {
  
    // Return a series of values from the given user account email based on the email parameter
    var regex = new RegExp(".+?(?=@)","gm"); // Use Regex to grab everything before the @ symbol from the user email
    var user_name = email.match(regex).toString(); // Get the match from the passed in string
    var user_name_split = user_name.split("."); // Use the "." as the delimiter for the first and last name of the user
    var first_name = user_name_split[0].charAt(0).toUpperCase() + user_name_split[0].slice(1); // Set the first letter of the First name to a Capital letter
    var last_name = user_name_split[1].charAt(0).toUpperCase() + user_name_split[1].slice(1); // set the First letter of the Last Name to a capital Letter
    var full_name = first_name + " " + last_name; // Write out the full name of the user as "Firstname Lastname"
  
    return [user_name, full_name];
  
  }
  
  function return_all_sheet_values(sheet) {
  
    // Return a range of values of all the sheet values
    var all_range = sheet.getDataRange(); // Let's get all the values in this sheet into a dataRange object
    var all_values = all_range.getValues(); // Convert the datarange object into actual readable values
  
    return all_values;
  
  }
  
  function is_specific_column(this_event, column) {
  
    // Return True if the edited cell matches the column of the given column
    var this_event_column = this_event.range.getColumn();
    if (this_event_column === column) {
      return true;
    }
    return false;
  
  }
  
  function is_row_value_matches(this_row_value, value) {
  
    // Return True if the row value passed in matches a given value
    if (this_row_value == value) {
      return true;
    }
    return false;
  
  }
  
  function is_there_duplicate_in_column(data_range_object, row_index, column, value) {
  
    // Return True if there is a duplicate for the give column
    var shipping_cell_number = "";
    for (n=1; n < data_range_object.length; ++n) {
      shipping_cell_number = data_range_object[n][column - 1];
      // As we loop through the Fedex column, let's check each value against the fedex tracking number, but make sure we don't check the number itself
      if ((shipping_cell_number === value) && (row_index !== (n + 1)) && (shipping_cell_number !== "")) { // We need to add one because the actual rowindex does not have a starting index of 0
        return true;
      }
    }
    return false;
  
  }
  
  function is_value_not_a_number(value) {
  
    // Return True if the passed in value is typeof == number
    if (typeof value !== "number" && value !== "") {
      return true;
    }
    return false;
  
  }
  
  function send_new_hire_setup_email(full_name, user_name, email, tracking_number, date) {
  
    // Send out the initial setup email to the specified new_hire and return a message stating that the email was sent
    var full_tracking_url = tracking_url + tracking_number;
    var add_bcc = bcc_email;
    var email_subject = "Your Laptop Is On The Way!";
    var fake_body = ''; // This is a blank because this part of the sendEmail method does not support html, we pass it later as part of keyvalue pair as real body.
    var real_body =  '<p>Hi <strong>' + full_name + '</strong>!<p>';
      real_body += '<p>We wanted to let you know that your computer has shipped for your start date: ' + date + '. You can get your tracking information from here: <a href="' + full_tracking_url + '">' + tracking_number + '</a>.<p>';
      real_body += '<p>Once you receive it, please set it up with the user account name: <strong>' + user_name + '</strong></p>';
      real_body += '<p>For an example, John Doe would be: <strong>john.doe</strong></p>';
      real_body += '<p>If you run into any issues logging into the computer please do not hesitate to contact us by responding to this email</p>';
      real_body += it_team_name;
      real_body += company_name;
  
    GmailApp.sendEmail(
    email,
    email_subject,
    fake_body,
    {
      bcc: add_bcc,
      htmlBody: real_body
    });
  
  }
  
  function send_reminder_setup_email(full_name, user_name, email, date) {
  
    // Send a reminder email to employees to finish up their computer setup.
  
    var add_bcc = bcc_email;
    var email_subject = "Reminder to setup your Laptop!";
    var fake_body = '';
    var real_body = '<p>Hi <strong>' + full_name + '!</strong></p>';
    real_body += '<p>The ' + it_team_name + ' at ' + company_name + ' is excited to meet you on ' + date + ' for your first day!</p>';
    real_body += '<p>This is a reminder to complete the computer setup prior to the session.</p>';
    real_body += '<br>';
    real_body += '<p>Please make sure that when you create your computer user account, it is formatted as <strong>' + user_name + '</strong> as it appears in your work email address.</p>';
    real_body += '<p>If you run into any issues please do not hesitate to reach out by replying to this email. See you soon!</p>';
    real_body += '<br>';
    real_body += it_team_name;
    real_body += company_name;
  
    GmailApp.sendEmail(
      email,
      email_subject,
      fake_body,
      {
        bcc: add_bcc,
        htmlBody: real_body
      });
  
  }
  
  function string_contains_value(string, value) {
  
    // Return True if a given string contains a given value
    if (string.toString().includes(value)) {
      return true;
    }
    return false;
  
  }
  
  function set_range_value(sheet, row, column, value) {
  
    // Set a specific value for a specific range_object
    sheet.getRange(row, column).setValue(value);
  
  }
  
  function is_value_empty(value) {
  
    // Briefly return True if the value passed in is empty, and if not return False
    if (! value) {
      return true;
    }
    return false;
  
  }
  
  function is_string_not_this_long(value, length) {
  
    // Return True if the passed in value is of a given length
    if (value.toString().length !== length && value !== "") {
      return true;
    }
    return false;
  
  }
  
  ///////////////////////////////////////////
  
  
  
  // Main function to run whenever an edit event happens
  function main(e) {
  
    var this_event = e;
  
    // Check if the current sheet being worked on is the Shipping Sheet
    if (is_current_spreadsheet_this_sheet("Shipping")) {
  
      ///////////////////////////////////////////////
      // Let's grab all the current available values from the sheet
      var current_sheet = return_current_sheet();
      let row_values = get_edited_row(this_event, current_sheet);
      const this_event_row_index = row_values[0];
      const this_event_row_list = row_values[1];
  
      var start_date = this_event_row_list[state_date_column - 1];
      var formatted_date = Utilities.formatDate(new Date(start_date), "GMT", "EEEEE, MMMMM d"); // Format the date so that it appears like example: "Thursday, January 1"
      var full_name = this_event_row_list[full_name_column - 1];
  
      var personal_email = this_event_row_list[personal_email_column - 1];
      var user_name = this_event_row_list[user_name_column - 1]
  
      var shipping_tracking_number = this_event_row_list[shipping_number_column - 1];
  
      var work_email = this_event_row_list[work_email_column - 1];
  
      var computer_shipped_status = this_event_row_list[computer_shipped_status_column - 1]; // This is the value inside the cell for the Computer Shipped Column
      var setup_email_sent_status = this_event_row_list[setup_email_sent_status_column - 1]; // This is the email status from any row from the Email Sent Column
  
      var to_send_reminder = this_event_row_list[send_reminder_column - 1];
      var sent_reminder_status = this_event_row_list[reminder_email_sent_status_column - 1];
  
      var sheets_values = return_all_sheet_values(current_sheet);
      ///////////////////////////////////////////////
  
  
  
      //// SET USER NAME ///////////////////////////////////
      // Set the full name and username columns if the edited cell belongs is the work email column
      if (is_specific_column(this_event, work_email_column)) {
  
        if (string_contains_value(work_email, company_domain)) {
          var user_name_from_work_email = ""
          var full_name_from_work_email = ""
          let user_name_values = get_username_info(work_email);
  
          user_name_from_work_email = user_name_values[0];
          full_name_from_work_email = user_name_values[1];
  
          set_range_value(current_sheet, this_event_row_index, full_name_column, full_name_from_work_email); // Set the full name cell
          set_range_value(current_sheet, this_event_row_index, user_name_column, user_name_from_work_email); // Set the user name cell
        } else if (is_value_empty(work_email)) {
          return;
        } else {
          ui.alert("To set the Username and Full name, please enter a work email that ends with " + company_domain + ". Example: john.doe" + company_domain + " will be john.doe");
        }
      }
      ////////////////////////////////////////////////////////
  
  
  
  
      //// SEND NEW HIRE SETUP EMAIL /////////////////////////
      // Send out an email to a given user if the edited cell belongs to the computer shipped status column and it is set to yes
      if (is_specific_column(this_event, computer_shipped_status_column)) {
  
        // Expect the trigger of Yes when a user wants to send an email.
        if (is_row_value_matches(computer_shipped_status, "Yes")) {
  
          // The following variables will store responses
          var no_start_date_response = "You have not set the start date. Please add a start date to send an email.";
          var no_full_name_response = "There is no Full Name set. Please check that you have set a Full Name.";
          var no_personal_email_response = "You have not added a personal email to reach out to the New Hire. Please add a personal email.";
          var no_user_name_response = "There is no user name set. Please add a user name to send the email.";
          var tracking_not_number_response = "The Fedex tracking number you inputted is not a number. Please verify that this is a number.";
          var no_tracking_number_response = "Please check that you have entered a tracking number. There is no tracking number set.";
          var tracking_number_nogood_length_response = "Please check the length of your Fedex Tracking Number. The length should be 12.";
          var a_wild_duplicate_appeared_response = "You have a duplicate with the value: " + shipping_tracking_number;
          var setup_email_sent_already_response = "ERROR. The setup email for " + full_name + " was already sent out. Delete the Email Sent Flag to resend.";
  
          var setup_email_sent_response = "The setup email was sent out for: " + full_name +  " to set up their computer with account name:\n\n" + user_name;
          var no_work_email_set_response = "You did not set a proper Work User Email. Are you sure you want to still send the tracking email?";
  
          // Create a quick list of checks that we can iterate through
          var checks = [
            {fx: is_value_empty, arguments: [start_date], response: no_start_date_response},
            {fx: is_value_empty, arguments: [full_name], response: no_full_name_response},
            {fx: is_value_empty, arguments: [personal_email], response: no_personal_email_response},
            {fx: is_value_empty, arguments: [user_name], response: no_user_name_response},
            {fx: is_value_not_a_number, arguments: [shipping_tracking_number], response: tracking_not_number_response},
            {fx: is_value_empty, arguments: [shipping_tracking_number], response: no_tracking_number_response},
            {fx: is_string_not_this_long, arguments: [shipping_tracking_number, 12], response: tracking_number_nogood_length_response},
            {fx: is_there_duplicate_in_column, arguments: [sheets_values, this_event_row_index, shipping_number_column, shipping_tracking_number], response: a_wild_duplicate_appeared_response},
            {fx: is_row_value_matches, arguments: [setup_email_sent_status, email_sent], response: setup_email_sent_already_response},
          ];
  
          var error_response = "";
          var is_there_error = false;
          checks.forEach(function(row) {
            if (row.fx.apply(null, row.arguments)) {
              error_response += "-- " + row.response + "\n";
              is_there_error = true;
            }
          })
  
          if (is_there_error) {
            ui.alert("Please correct the following in order to send an email:\n\n\n " + error_response);
            return;
          }
  
          // Let's check if we have a work email sent. We can still send the email if not but we want to verify that the user is ok with no work email
          // Otherwise blast that email out!
          if (is_value_empty(work_email)) {
            var does_user_want_to_send_email_still = ui.alert(no_work_email_set_response, ui.ButtonSet.YES_NO);
            if (does_user_want_to_send_email_still == ui.Button.NO) {
              return;
            }
          }
  
          // send email for new hire setup since there is no error
          send_new_hire_setup_email(full_name, user_name, personal_email, shipping_tracking_number, formatted_date);
          // Set the status of the email sent column to sent and prompt user that the email was sent.
          current_sheet.getRange(this_event_row_index, setup_email_sent_status_column).setValue(email_sent);
          SpreadsheetApp.flush();
          ui.alert(setup_email_sent_response);
  
  
        }
  
      }
  
      ////////////////////////////////////////////////////////
  
      //// SEND REMINDER SETUP EMAIL /////////////////////////
      // Send out a reminder to the new hire that they need to setup before their start date.
      if (is_specific_column(this_event, send_reminder_column)) {
  
        // Check first to make sure that the user set the Send reminder email to Yes and that the initial email was already sent
        if (is_row_value_matches(to_send_reminder, "Yes")) {
  
          // Prepare responses in variables for the user
          var reminder_setup_email_sent_response = "A follow up reminder email was sent out for: " + full_name + ". You can resend by deleting the Reminder Sent flag.";
          var reminder_email_already_sent_response = "ERROR. The follow up reminder email was already sent out: " + full_name + ". To resend, delete the Reminder Sent flag.";
          var setup_email_not_sent_yet_response = "ERROR. The initial setup email was not sent out yet. Please send out this email first before sending the reminder.";
  
          // Tell the user to first send the initial tracking setup email to send the reminder
          if (! is_row_value_matches(setup_email_sent_status, email_sent)) {
            ui.alert(setup_email_not_sent_yet_response);
            return;
          }
  
          // Tell the user the reminder was already sent
          if (is_row_value_matches(sent_reminder_status, reminder_sent)) {
            ui.alert(reminder_email_already_sent_response);
            return;
          }
  
          // Send the reminder email to the new hire that they need to finish their setup
          send_reminder_setup_email(full_name, user_name, personal_email, formatted_date);
          // Set the status of the reminder email sent column to Reminder sent and let the user know that the email was sent
          current_sheet.getRange(this_event_row_index, reminder_email_sent_status_column).setValue(reminder_sent);
          SpreadsheetApp.flush();
          ui.alert(reminder_setup_email_sent_response)
  
        }
      }
      ////////////////////////////////////////////////////////
  
  
    }
  
}