# User Email Lookup

The User Email Lookup tool lets you paste a list of display names and retrieve primary email addresses from Microsoft Graph. It reuses the shared workspace shell so you can quickly process long rosters without leaving the browser.

## Key Capabilities
- **Bulk lookups** – Paste any number of names (one per line) and fetch Microsoft 365 profile emails in sequence.
- **Smart matching** – Prefer exact display name matches when available, while still surfacing the closest match when only partial results exist.
- **Duplicate trimming** – Optionally remove repeated names before issuing Microsoft Graph requests to save time and avoid throttling.
- **Progress feedback** – Inline status updates show which record is being processed, and a results table summarizes matches, misses, and errors.
- **Export friendly** – Copy the results to the clipboard in a tab-delimited format or download them as a CSV for sharing.

## Usage
1. Grab a Microsoft Graph access token from the [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and paste it into the token panel.
2. Paste your list of user display names into the textarea—one per line.
3. Adjust lookup options if necessary (toggle duplicate removal, change the throttle delay, etc.).
4. Click **Lookup Emails**. Progress updates will appear in the setup panel and the table on the right will populate with results.
5. Copy or download the table once the lookup completes.

The tool calls `https://graph.microsoft.com/v1.0/users` with `$search` and `startswith` queries and requires a token with at least the `User.ReadBasic.All` permission.
