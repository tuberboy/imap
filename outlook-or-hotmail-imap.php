<?php
set_time_limit(0);

// IMAP configuration
// $hostname = '{imap-mail.outlook.com:993/imap/ssl/debug}INBOX'; // also works it
$hostname = '{outlook.office365.com:993/imap/ssl}INBOX'; // IMAP server and mailbox
$username = ''; // Email username or full email
$password = ''; // Email password

// Connect to the IMAP server
$inbox = imap_open($hostname, $username, $password) or die('Cannot connect to Outlook: ' . imap_last_error());

// Function to print a single email message
function printMessage($inbox, $email_number) {
    // Fetch email overview
    $overview = imap_fetch_overview($inbox, $email_number, 0);
    // Extract sender email, subject, and body from the email overview
    $sender_email = $overview[0]->from ?? 'N/A';
    $subject = $overview[0]->subject ?? 'N/A';
    $body = quoted_printable_decode(imap_fetchbody($inbox, $email_number, 1));
    // Print email details
    echo "Sender Email: $sender_email<br>";
    echo "Subject: $subject<br>";
    echo "Body: $body<br><br>";
}

// Mailbox Status
$mailboxStatus = imap_status($inbox, $hostname, SA_ALL);
if ($mailboxStatus !== false) {
    // Print mailbox status details
    echo "Messages: " . ($mailboxStatus->messages ?? 'N/A') . "<br>";
    echo "Recent: " . ($mailboxStatus->recent ?? 'N/A') . "<br>";
    echo "Unseen: " . ($mailboxStatus->unseen ?? 'N/A') . "<br>";
} else {
    // Print error message if failed to retrieve mailbox status
    echo "Failed to retrieve mailbox status.<br>";
}

// Number of Messages
$numMessages = imap_num_msg($inbox);
echo "Number of messages: " . ($numMessages !== false ? $numMessages : 'N/A') . "<br>";

// Define search criteria and corresponding labels
$searches = [
    'UNSEEN' => 'Search Results',
    'ALL' => 'All Messages',
    'ANSWERED' => 'Answered Messages',
    'DELETED' => 'Deleted Messages',
    'FLAGGED' => 'Flagged Messages',
    'UNFLAGGED' => 'Unflagged Messages',
    'SINCE "10-Feb-2022"' => "Messages received since 10th February 2022",
    'FROM "sender@example.com"' => "Messages from sender@example.com",
    'SUBJECT "important"' => "Messages with 'important' in the subject"
];

// Array to store unique message IDs
$uniqueMessageIDs = [];

// Iterate over each search criteria and fetch corresponding search results
foreach ($searches as $criteria => $label) {
    $searchResults = imap_search($inbox, $criteria);
    if ($searchResults !== false) {
        // Print label for search results
        echo "$label:<br>";
        // Iterate over search results and print messages
        foreach ($searchResults as $email_number) {
            // Check if the message has not been printed already
            if (!in_array($email_number, $uniqueMessageIDs)) {
                // Add message ID to the list of processed messages
                $uniqueMessageIDs[] = $email_number;
                // Print the message
                printMessage($inbox, $email_number);
            }
        }
    } else {
        // Print message if no search results found for the criteria
        echo "No $label found.<br>";
    }
}

// IMAP Errors
$imapErrors = imap_errors();
if ($imapErrors !== false) {
    // Print IMAP errors
    echo "IMAP Errors:<br>";
    print_r($imapErrors);
    echo "<br>";
}

// Close the IMAP connection
imap_close($inbox);
?>
