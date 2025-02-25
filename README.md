# Slack to PowerPoint Profile Generator

This script fetches users from a specified Slack channel and creates a PowerPoint presentation with a slide for each person including their profile picture and information.

## Features

- Fetches user profiles from a Slack channel
- Downloads high-resolution profile pictures
- Creates a professional PowerPoint slide for each team member
- Includes user details such as title, email, phone, and timezone

## Requirements

- Node.js (v14 or higher)
- Slack API token with appropriate permissions
- A Slack channel ID

## Setup

1. Clone or download this repository
2. Install dependencies:

```bash
npm install
```

3. Create a `.env` file by copying the example:

```bash
cp .env.example .env
```

4. Edit the `.env` file and add your Slack API token and channel ID:

```
SLACK_TOKEN=xoxb-your-token-here
SLACK_CHANNEL_ID=C12345678
```

## Getting a Slack API Token

1. Go to [https://api.slack.com/apps](https://api.slack.com/apps)
2. Click "Create New App" and select "From scratch"
3. Name your app and select your workspace
4. In the left sidebar, click on "OAuth & Permissions"
5. Scroll down to "Scopes" and add the following Bot Token Scopes:
   - `users:read`
   - `channels:read`
   - `groups:read`
   - `mpim:read`
   - `im:read`
6. Scroll up and click "Install to Workspace"
7. Copy the "Bot User OAuth Token" that starts with `xoxb-`

## Finding Your Channel ID

1. Open Slack in a web browser
2. Navigate to the channel you want to use
3. The channel ID is in the URL: `https://app.slack.com/client/T12345678/C87654321`
   (In this example, `C87654321` is the channel ID)

## Usage

First, validate your configuration:

```bash
npm run validate
```

This will check if your Slack token and channel ID are valid and accessible.

Once validation passes, run the main script:

```bash
npm start
```

The script will:
1. Connect to the Slack API
2. Fetch all users from the specified channel
3. Download their profile pictures to a temporary folder
4. Generate a PowerPoint presentation with a slide for each user
5. Save the presentation as `team_directory.pptx` in the project folder
6. Clean up temporary files

## Output

The script will create a file named `team_directory.pptx` in the root directory of the project. Open this file with Microsoft PowerPoint or a compatible application to view and edit the presentation.

## Troubleshooting

- **Error: SLACK_TOKEN is required:** Make sure you've created a `.env` file with your Slack token.
- **Error: SLACK_CHANNEL_ID is required:** Ensure you've added the channel ID to your `.env` file.
- **Error getting channel members:** Verify that your token has the correct permissions and the channel ID is valid.
- **No users found in the channel:** Confirm that the channel contains users and that your bot has been added to the channel.

## Limitations

- The script can only access channels that the bot has been invited to
- Rate limits may apply when fetching large numbers of users
- Images are temporarily stored during processing