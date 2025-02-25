import { WebClient } from '@slack/web-api';
import pptxgen from 'pptxgenjs';
import axios from 'axios';
import fs from 'fs/promises';
import path from 'path';
import dotenv from 'dotenv';
import { createWriteStream } from 'fs';
import { pipeline } from 'stream/promises';

// Load environment variables
dotenv.config();

// Initialize Slack API client
const slackToken = process.env.SLACK_TOKEN;
if (!slackToken) {
  console.error('Error: SLACK_TOKEN is required in .env file');
  process.exit(1);
}

const slack = new WebClient(slackToken);

// Create temp directory for profile images
const TEMP_DIR = path.join(process.cwd(), 'temp_images');

async function ensureTempDir() {
  try {
    await fs.mkdir(TEMP_DIR, { recursive: true });
    console.log(`Created temp directory at ${TEMP_DIR}`);
  } catch (error) {
    console.error('Error creating temp directory:', error);
    throw error;
  }
}

// Download profile image
async function downloadProfileImage(imageUrl, userId) {
  if (!imageUrl) {
    console.log(`No profile image URL for user ${userId}`);
    return null;
  }
  
  // Use 512px version of the image if available
  const highResImageUrl = imageUrl.replace(/\d+$/, '512');
  const imagePath = path.join(TEMP_DIR, `${userId}.jpg`);
  
  try {
    const response = await axios({
      method: 'get',
      url: highResImageUrl,
      responseType: 'stream'
    });
    
    await pipeline(response.data, createWriteStream(imagePath));
    console.log(`Downloaded profile image for ${userId}`);
    return imagePath;
  } catch (error) {
    console.error(`Error downloading profile image for ${userId}:`, error.message);
    return null;
  }
}

// Get all users from a channel
async function getUsersInChannel(channelId) {
  try {
    // Get all members in the channel
    const result = await slack.conversations.members({
      channel: channelId
    });
    
    const memberIds = result.members;
    console.log(`Found ${memberIds.length} users in channel`);
    
    // Fetch detailed info for each user
    const userDetails = [];
    
    for (const userId of memberIds) {
      try {
        const userInfo = await slack.users.info({ user: userId });
        
        if (userInfo.user && !userInfo.user.is_bot && !userInfo.user.deleted) {
          userDetails.push(userInfo.user);
        }
      } catch (error) {
        console.error(`Error fetching info for user ${userId}:`, error.message);
      }
    }
    
    console.log(`Retrieved details for ${userDetails.length} users`);
    return userDetails;
  } catch (error) {
    console.error('Error getting channel members:', error.message);
    throw error;
  }
}

// Create PowerPoint slides
async function createPowerPoint(users) {
  const pres = new pptxgen();
  
  // Set default slide size to 16:9
  pres.layout = 'LAYOUT_16x9';
  
  // Create title slide
  const titleSlide = pres.addSlide();
  titleSlide.addText("Team Directory", {
    x: 1.0,
    y: 2.0,
    w: '80%',
    h: 1.5,
    fontSize: 44,
    color: '363636',
    fontFace: 'Arial',
    align: 'center',
    bold: true
  });
  
  // Create a slide for each user
  for (const user of users) {
    try {
      // Download profile image
      const imagePath = await downloadProfileImage(
        user.profile.image_original || user.profile.image_512 || user.profile.image_192,
        user.id
      );
      
      // Create a new slide
      const slide = pres.addSlide();
      
      // Add user name as title
      slide.addText(user.profile.real_name || user.name, {
        x: 0.5,
        y: 0.5,
        w: '90%',
        h: 1.0,
        fontSize: 36,
        color: '363636',
        fontFace: 'Arial',
        bold: true
      });
      
      // Add profile image if available
      if (imagePath) {
        slide.addImage({
          path: imagePath,
          x: 0.5,
          y: 1.7,
          w: 3.0,
          h: 3.0
        });
      }
      
      // Add profile info
      const profileInfo = [];
      
      if (user.profile.title) profileInfo.push(`Title: ${user.profile.title}`);
      if (user.profile.email) profileInfo.push(`Email: ${user.profile.email}`);
      if (user.profile.phone) profileInfo.push(`Phone: ${user.profile.phone}`);
      if (user.tz) profileInfo.push(`Timezone: ${user.tz}`);
      
      slide.addText(profileInfo, {
        x: 4.0,
        y: 1.7,
        w: 5.5,
        h: 3.0,
        fontSize: 18,
        color: '666666',
        fontFace: 'Arial',
        bullet: true,
        lineSpacing: 30
      });
      
      console.log(`Created slide for ${user.profile.real_name || user.name}`);
    } catch (error) {
      console.error(`Error creating slide for user ${user.id}:`, error.message);
    }
  }
  
  // Save the presentation
  const outputPath = path.join(process.cwd(), 'team_directory.pptx');
  await pres.writeFile({ fileName: outputPath });
  console.log(`PowerPoint saved to ${outputPath}`);
  
  return outputPath;
}

// Clean up temp directory
async function cleanupTempDir() {
  try {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    console.log('Cleaned up temporary files');
  } catch (error) {
    console.error('Error cleaning up temp directory:', error);
  }
}

// Main function
async function main() {
  const channelId = process.env.SLACK_CHANNEL_ID;
  
  if (!channelId) {
    console.error('Error: SLACK_CHANNEL_ID is required in .env file');
    process.exit(1);
  }
  
  try {
    await ensureTempDir();
    const users = await getUsersInChannel(channelId);
    
    if (users.length === 0) {
      console.log('No users found in the channel');
      return;
    }
    
    const presentationPath = await createPowerPoint(users);
    console.log(`Success! Presentation created at: ${presentationPath}`);
    
    // Clean up temp files
    await cleanupTempDir();
  } catch (error) {
    console.error('Error:', error.message);
    process.exit(1);
  }
}

// Run the script
main();