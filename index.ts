import { google } from "googleapis";
import { config } from "dotenv";
import * as xlsx from "xlsx";

config(); // Load environment variables

// Load API Keys from .env file
const API_KEY = process.env.GOOGLE_API_KEY as string;
const GOOGLE_SHEETS_ID_INPUT = process.env.GOOGLE_SHEETS_ID_INPUT as string;
const GOOGLE_SHEETS_ID_OUTPUT = process.env.GOOGLE_SHEETS_ID_OUTPUT as string;
const GOOGLE_APPLICATION_CREDENTIALS = process.env
  .GOOGLE_APPLICATION_CREDENTIALS as string;

if (!API_KEY)
  throw new Error(
    "YOUTUBE_API_KEY is missing. Please set it in your .env file."
  );
if (
  !GOOGLE_SHEETS_ID_INPUT ||
  !GOOGLE_SHEETS_ID_OUTPUT ||
  !GOOGLE_APPLICATION_CREDENTIALS
)
  throw new Error("Google Sheets file IDs are missing in .env.");

// Google API Authentication
const auth = new google.auth.GoogleAuth({
  keyFile: GOOGLE_APPLICATION_CREDENTIALS,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ],
});
const sheets = google.sheets({ version: "v4", auth });

// Interfaces for API Responses
interface YouTubeApiResponse<T> {
  items: T[];
}

interface Channel {
  id: string;
}

interface VideoSnippet {
  title: string;
  channelId: string;
  description: string;
  channelTitle: string;
  publishedAt: Date | string;
  publishTime: Date | string;
}

interface VideoId {
  videoId: string;
}

interface VideoItem {
  id: VideoId;
  snippet: VideoSnippet;
}

interface VideoData {
  title: string;
  url: string;
  channelId: string;
  description: string;
  channelTitle: string;
  publishedAt: Date | string;
  publishTime: Date | string;
}

/**
 * Reads a list of YouTube URLs from a Google Sheet.
 * @param sheetId Google Sheets ID
 * @returns An array of YouTube URLs
 */
async function readGoogleSheet(sheetId: string): Promise<string[]> {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "C:C", // Assuming URLs are in column A
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return [];

    // Regex to match YouTube URLs
    const youtubeRegex =
      /(https?:\/\/)?(www\.)?(youtube\.com|youtu\.be)\/[^\s]+/gi;

    return rows
      .flat()
      .map((url) => url.trim())
      .filter((url) => youtubeRegex.test(url)); // Keep only YouTube URLs
  } catch (error) {
    console.error("Error reading Google Sheets:", error);
    return [];
  }
}

/**
 * Extracts the YouTube handle from a given channel URL.
 * @param url YouTube channel URL
 * @returns Extracted YouTube handle or null if invalid
 */
function extractHandleFromUrl(url: string): string | null {
  const match = url.match(/youtube\.com\/@([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Fetches the channel ID from a YouTube handle.
 * @param handle YouTube handle (e.g., "@LinusTechTips")
 * @returns The channel ID or null if not found
 */
async function getChannelId(handle: string): Promise<string | null> {
  const url = `https://www.googleapis.com/youtube/v3/channels?part=id&forHandle=${handle}&key=${API_KEY}`;
  console.log(url);
  try {
    const response = await fetch(url);
    const json = (await response.json()) as YouTubeApiResponse<Channel>;

    return json.items?.[0]?.id || null;
  } catch (error) {
    console.error(`Error fetching channel ID for ${handle}:`, error);
    return null;
  }
}

/**
 * Fetches the latest videos from a given YouTube channel ID.
 * @param channelId YouTube channel ID
 * @returns An array of video data
 */
async function getVideos(channelId: string): Promise<VideoData[]> {
  const url = `https://www.googleapis.com/youtube/v3/search?key=${API_KEY}&channelId=${channelId}&part=snippet&type=video&maxResults=10`;

  try {
    const response = await fetch(url);
    const json = (await response.json()) as YouTubeApiResponse<VideoItem>;
    console.log(json);
    return json.items.map((video) => ({
      title: video.snippet.title,
      url: `https://www.youtube.com/watch?v=${video.id.videoId}`,
      channelId: video.snippet.channelId,
      description: video.snippet.description,
      channelTitle: video.snippet.channelTitle,
      publishedAt: video.snippet.publishedAt,
      publishTime: video.snippet.publishTime,
    }));
  } catch (error) {
    console.error(`Error fetching videos for channel ${channelId}:`, error);
    return [];
  }
}

/**
 * Writes extracted video data to a Google Sheet.
 * @param sheetId Google Sheets ID
 * @param data Array of video data to write to the sheet
 */
async function writeToGoogleSheet(
  sheetId: string,
  data: VideoData[]
): Promise<void> {
  try {
    const values = Object.values(data) as any[];

    await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: "A1", // Starting from the first row
      valueInputOption: "RAW",
      requestBody: {
        values,
      },
    });
    console.log("Data successfully written to Google Sheets.");
  } catch (error) {
    console.error("Error writing to Google Sheets:", error);
  }
}

/**
 * Writes extracted video data to a local Excel file.
 * @param filePath Path to the Excel file
 * @param data Array of video data to write
 */
async function writeToExcel(
  filePath: string,
  data: VideoData[]
): Promise<void> {
  try {
    if (data.length === 0) {
      console.log("No video data to save.");
      return;
    }

    // Prepare the worksheet data
    console.log(data, Object.keys(data), Object.values(data));
    const headers = Object.keys(data).map((el) => el.toUpperCase()) as any[]; // Column headers
    const values = Object.values(data) as any[];

    // Create a new workbook and worksheet
    const worksheet = xlsx.utils.aoa_to_sheet([...headers, ...values]);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, "Videos");

    // Write to file
    xlsx.writeFile(workbook, filePath);
    console.log(`Data successfully written to ${filePath}`);
  } catch (error) {
    console.error("Error writing to Excel file:", error);
  }
}

/**
 * Main function: Reads YouTube URLs, fetches video details, and writes results.
 */
async function processYouTubeChannels(): Promise<void> {
  const urls: string[] = (await readGoogleSheet(GOOGLE_SHEETS_ID_INPUT)).slice(
    0,
    5
  );

  if (urls.length === 0) {
    console.log("No YouTube URLs found in the input sheet.");
    return;
  }

  const allVideos: VideoData[] = [];

  for (const url of urls) {
    const handle: string | null = extractHandleFromUrl(url);
    if (!handle) {
      console.error(`Invalid YouTube URL: ${url}`);
      continue;
    }

    const channelId: string | null = await getChannelId(handle);
    if (!channelId) {
      console.error(`Channel ID not found for ${handle}`);
      continue;
    }

    const videos: VideoData[] = await getVideos(channelId);
    allVideos.push(...videos);
  }

  // await writeToGoogleSheet(GOOGLE_SHEETS_ID_OUTPUT, allVideos);

  writeToExcel("videos.xlsx", allVideos);
}

// Run the function
processYouTubeChannels();
