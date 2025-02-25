import { WebClient } from '@slack/web-api';
import dotenv from 'dotenv';
import fs from 'fs';
import path from 'path';

// Load environment variables
dotenv.config();

// Check if .env file exists
function checkEnvFile() {
  const envPath = path.join(process.cwd(), '.env');
  
  if (!fs.existsSync(envPath)) {
    console.error('\x1b[31m✗ .env file not found!\x1b[0m');
    console.log('  Please create a .env file by copying .env.example and filling in your values.');
    return false;
  }
  
  console.log('\x1b[32m✓ .env file exists\x1b[0m');
  return true;
}

// Check required environment variables
function checkEnvironmentVariables() {
  let missingVars = [];
  
  if (!process.env.SLACK_TOKEN) missingVars.push('SLACK_TOKEN');
  if (!process.env.SLACK_CHANNEL_ID) missingVars.push('SLACK_CHANNEL_ID');
  
  if (missingVars.length > 0) {
    console.error(`\x1b[31m✗ Missing required environment variables: ${missingVars.join(', ')}\x1b[0m`);
    return false;
  }
  
  console.log('\x1b[32m✓ All required environment variables are present\x1b[0m');
  return true;
}

// Validate Slack token
async function validateSlackToken() {
  const slack = new WebClient(process.env.SLACK_TOKEN);
  
  try {
    const authTest = await slack.auth.test();
    console.log('\x1b[32m✓ Slack token is valid\x1b[0m');
    console.log(`  Connected as: ${authTest.user} (${authTest.user_id})`);
    console.log(`  Workspace: ${authTest.team} (${authTest.team_id})`);
    return true;
  } catch (error) {
    console.error('\x1b[31m✗ Slack token validation failed!\x1b[0m');
    console.error(`  Error: ${error.message}`);
    return false;
  }
}

// Validate channel ID
async function validateChannelId() {
  const slack = new WebClient(process.env.SLACK_TOKEN);
  
  try {
    const channelInfo = await slack.conversations.info({
      channel: process.env.SLACK_CHANNEL_ID
    });
    
    console.log('\x1b[32m✓ Channel ID is valid\x1b[0m');
    console.log(`  Channel: ${channelInfo.channel.name} (${channelInfo.channel.id})`);
    
    // Check channel membership count
    const membersResult = await slack.conversations.members({
      channel: process.env.SLACK_CHANNEL_ID
    });
    
    console.log(`  Members: ${membersResult.members.length}`);
    
    return true;
  } catch (error) {
    console.error('\x1b[31m✗ Channel ID validation failed!\x1b[0m');
    console.error(`  Error: ${error.message}`);
    console.log('  Note: Make sure the bot is a member of the channel. For private channels, you need to invite the bot.');
    return false;
  }
}

// Main validation function
async function validate() {
  console.log('Validating Slack to PowerPoint configuration...\n');
  
  const envFileExists = checkEnvFile();
  if (!envFileExists) return false;
  
  const varsExist = checkEnvironmentVariables();
  if (!varsExist) return false;
  
  const tokenValid = await validateSlackToken();
  if (!tokenValid) return false;
  
  const channelValid = await validateChannelId();
  if (!channelValid) return false;
  
  console.log('\n\x1b[32m✓ All validation checks passed! Your configuration is valid.\x1b[0m');
  console.log('  You can now run the main script: npm start');
  return true;
}

// Run validation
validate().catch(error => {
  console.error('\n\x1b[31mUnexpected error during validation:\x1b[0m', error);
  process.exit(1);
});