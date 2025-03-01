import { google } from "googleapis";
import { config } from "dotenv";
import * as xlsx from "xlsx";

config();

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

const auth = new google.auth.GoogleAuth({
  keyFile: GOOGLE_APPLICATION_CREDENTIALS,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ],
});
const sheets = google.sheets({ version: "v4", auth });

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
  viewCount?: string;
  likeCount?: string;
  commentCount?: string;
  duration?: string;
  readableDuration?: string;
}

async function readGoogleSheet(sheetId: string): Promise<string[]> {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "C:C",
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return [];

    const youtubeRegex =
      /(https?:\/\/)?(www\.)?(youtube\.com|youtu\.be)\/[^\s]+/gi;

    return rows
      .flat()
      .map((url) => url.trim())
      .filter((url) => youtubeRegex.test(url));
  } catch (error) {
    console.error("Error reading Google Sheets:", error);
    return [];
  }
}

function extractHandleFromUrl(url: string): string | null {
  const match = url.match(/youtube\.com\/@([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

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

async function getVideos(channelId: string): Promise<VideoData[]> {
  const url = `https://www.googleapis.com/youtube/v3/search?key=${API_KEY}&channelId=${channelId}&part=snippet&type=video&maxResults=10`;

  try {
    const response = await fetch(url);
    const json = (await response.json()) as YouTubeApiResponse<VideoItem>;

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

async function getVideoDetails(videoId: string): Promise<any> {
  const url = `https://www.googleapis.com/youtube/v3/videos?part=statistics,contentDetails&id=${videoId}&key=${API_KEY}`;

  try {
    const response = await fetch(url);
    const json = await response.json();
    if (json.items && json.items.length > 0) {
      return json.items[0];
    }
    return null;
  } catch (error) {
    console.error(`Error fetching video details for ${videoId}:`, error);
    return null;
  }
}

function parseDuration(duration: string): string {
  const match = duration.match(/PT(\d+H)?(\d+M)?(\d+S)?/);
  if (!match) return duration;

  const hours = parseInt(match[1]?.replace("H", "") || "0", 10);
  const minutes = parseInt(match[2]?.replace("M", "") || "0", 10);
  const seconds = parseInt(match[3]?.replace("S", "") || "0", 10);

  let readableDuration = "";
  if (hours > 0) readableDuration += `${hours}h `;
  if (minutes > 0) readableDuration += `${minutes}m `;
  if (seconds > 0) readableDuration += `${seconds}s`;

  return readableDuration.trim();
}

async function writeToExcel(
  filePath: string,
  data: VideoData[]
): Promise<void> {
  try {
    if (data.length === 0) {
      console.log("No video data to save.");
      return;
    }

    const headers = Object.keys(data[0]);
    const values = data.map((obj) => Object.values(obj));

    const worksheet = xlsx.utils.aoa_to_sheet([headers, ...values]);

    // Set column widths
    const columnWidths = headers.map(() => ({ wch: 30 }));
    worksheet["!cols"] = columnWidths;

    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, "Videos");

    xlsx.writeFile(workbook, filePath);
    console.log(`Data successfully written to ${filePath}`);
  } catch (error) {
    console.error("Error writing to Excel file:", error);
  }
}

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

    for (const video of videos) {
      const videoId = video.url.split("v=")[1];
      const details = await getVideoDetails(videoId);
      if (details) {
        video.viewCount = details.statistics.viewCount;
        video.likeCount = details.statistics.likeCount;
        video.commentCount = details.statistics.commentCount;
        // video.duration = details.contentDetails.duration;
        video.readableDuration = parseDuration(video.duration || "");
      }
    }
    allVideos.push(...videos);
  }
  writeToExcel("videos.xlsx", allVideos);
}

processYouTubeChannels();
