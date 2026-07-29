import { Client, Databases } from "appwrite";

const endpoint = import.meta.env.VITE_APPWRITE_ENDPOINT;
const projectId = import.meta.env.VITE_APPWRITE_PROJECT_ID;

export const DATABASE_ID = import.meta.env.VITE_APPWRITE_DATABASE_ID;
export const TABLE_ID = import.meta.env.VITE_APPWRITE_TABLE_ID;

if (!endpoint || !projectId || !DATABASE_ID || !TABLE_ID) {
  console.error(
    "Missing one of VITE_APPWRITE_ENDPOINT / VITE_APPWRITE_PROJECT_ID / VITE_APPWRITE_DATABASE_ID / VITE_APPWRITE_TABLE_ID. Set them in a .env.local file (local dev) or in your Vercel project's Environment Variables (production)."
  );
}

const client = new Client().setEndpoint(endpoint).setProject(projectId);

export const databases = new Databases(client);
