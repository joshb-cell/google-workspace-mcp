import OAuthProvider from "@cloudflare/workers-oauth-provider";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { McpAgent } from "agents/mcp";
import { z } from "zod";
import { GoogleHandler } from "./google-handler";
import { googleApiFetch, type Props } from "./utils";

export class MyMCP extends McpAgent<Env, Record<string, never>, Props> {
	server = new McpServer({
		name: "Google Workspace MCP Server",
		version: "1.0.0",
	});

	async init() {
		// ===== DRIVE TOOLS =====

		this.server.tool(
			"drive_list_files",
			"List or search files in Google Drive",
			{
				query: z.string().optional().describe("Search query (Google Drive search syntax)"),
				pageSize: z.number().default(20).describe("Number of results to return"),
				pageToken: z.string().optional().describe("Token for next page of results"),
			},
			async ({ query, pageSize, pageToken }) => {
				const params = new URLSearchParams({
					fields: "files(id,name,mimeType,modifiedTime,size,parents,webViewLink),nextPageToken",
					pageSize: String(pageSize),
				});
				if (query) params.set("q", query);
				if (pageToken) params.set("pageToken", pageToken);

				const resp = await googleApiFetch(
					`https://www.googleapis.com/drive/v3/files?${params}`,
					this.props!.googleAccessToken,
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error listing files: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
			},
		);

		this.server.tool(
			"drive_read_file",
			"Read the content of a file from Google Drive. Automatically exports Google Docs/Sheets/Slides to text formats.",
			{
				fileId: z.string().describe("The ID of the file to read"),
			},
			async ({ fileId }) => {
				// First get metadata to determine type
				const metaResp = await googleApiFetch(
					`https://www.googleapis.com/drive/v3/files/${fileId}?fields=mimeType,name`,
					this.props!.googleAccessToken,
				);

				if (!metaResp.ok) {
					const error = await metaResp.text();
					return { content: [{ type: "text", text: `Error reading file metadata: ${error}` }], isError: true };
				}

				const meta = (await metaResp.json()) as { mimeType: string; name: string };

				// Google native formats need export
				const exportMap: Record<string, string> = {
					"application/vnd.google-apps.document": "text/plain",
					"application/vnd.google-apps.spreadsheet": "text/csv",
					"application/vnd.google-apps.presentation": "text/plain",
				};

				let contentResp: Response;
				if (exportMap[meta.mimeType]) {
					contentResp = await googleApiFetch(
						`https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=${encodeURIComponent(exportMap[meta.mimeType])}`,
						this.props!.googleAccessToken,
					);
				} else {
					contentResp = await googleApiFetch(
						`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
						this.props!.googleAccessToken,
					);
				}

				if (!contentResp.ok) {
					const error = await contentResp.text();
					return { content: [{ type: "text", text: `Error reading file content: ${error}` }], isError: true };
				}

				const content = await contentResp.text();
				return { content: [{ type: "text", text: `File: ${meta.name}\n\n${content}` }] };
			},
		);

		this.server.tool(
			"drive_create_file",
			"Create a new file in Google Drive",
			{
				name: z.string().describe("File name"),
				content: z.string().describe("File content"),
				mimeType: z.string().default("text/plain").describe("MIME type of the content"),
				parentFolderId: z.string().optional().describe("Parent folder ID"),
			},
			async ({ name, content, mimeType, parentFolderId }) => {
				const metadata: any = { name, mimeType };
				if (parentFolderId) metadata.parents = [parentFolderId];

				const boundary = "boundary_" + Date.now();
				const body =
					`--${boundary}\r\n` +
					`Content-Type: application/json; charset=UTF-8\r\n\r\n` +
					`${JSON.stringify(metadata)}\r\n` +
					`--${boundary}\r\n` +
					`Content-Type: ${mimeType}\r\n\r\n` +
					`${content}\r\n` +
					`--${boundary}--`;

				const resp = await fetch(
					"https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink",
					{
						method: "POST",
						headers: {
							Authorization: `Bearer ${this.props!.googleAccessToken}`,
							"Content-Type": `multipart/related; boundary=${boundary}`,
						},
						body,
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error creating file: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: `File created: ${JSON.stringify(data, null, 2)}` }] };
			},
		);

		this.server.tool(
			"drive_create_folder",
			"Create a new folder in Google Drive",
			{
				name: z.string().describe("Folder name"),
				parentFolderId: z.string().optional().describe("Parent folder ID"),
			},
			async ({ name, parentFolderId }) => {
				const metadata: any = {
					name,
					mimeType: "application/vnd.google-apps.folder",
				};
				if (parentFolderId) metadata.parents = [parentFolderId];

				const resp = await googleApiFetch(
					"https://www.googleapis.com/drive/v3/files?fields=id,name,webViewLink",
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify(metadata),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error creating folder: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: `Folder created: ${JSON.stringify(data, null, 2)}` }] };
			},
		);

		this.server.tool(
			"drive_delete_file",
			"Delete a file or folder from Google Drive",
			{
				fileId: z.string().describe("The ID of the file or folder to delete"),
			},
			async ({ fileId }) => {
				const resp = await googleApiFetch(
					`https://www.googleapis.com/drive/v3/files/${fileId}`,
					this.props!.googleAccessToken,
					{ method: "DELETE" },
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error deleting file: ${error}` }], isError: true };
				}

				return { content: [{ type: "text", text: "File deleted successfully." }] };
			},
		);

		this.server.tool(
			"drive_move_file",
			"Move a file to a different folder in Google Drive",
			{
				fileId: z.string().describe("The ID of the file to move"),
				newParentFolderId: z.string().describe("The ID of the destination folder"),
			},
			async ({ fileId, newParentFolderId }) => {
				// Get current parents
				const metaResp = await googleApiFetch(
					`https://www.googleapis.com/drive/v3/files/${fileId}?fields=parents`,
					this.props!.googleAccessToken,
				);

				if (!metaResp.ok) {
					const error = await metaResp.text();
					return { content: [{ type: "text", text: `Error getting file info: ${error}` }], isError: true };
				}

				const meta = (await metaResp.json()) as { parents?: string[] };
				const oldParents = (meta.parents || []).join(",");

				const resp = await googleApiFetch(
					`https://www.googleapis.com/drive/v3/files/${fileId}?addParents=${newParentFolderId}&removeParents=${oldParents}&fields=id,name,parents`,
					this.props!.googleAccessToken,
					{ method: "PATCH", body: JSON.stringify({}) },
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error moving file: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: `File moved: ${JSON.stringify(data, null, 2)}` }] };
			},
		);

		this.server.tool(
			"drive_share_file",
			"Share a file or folder with another person via email. Can set role to reader, commenter, or writer.",
			{
				fileId: z.string().describe("The ID of the file or folder to share"),
				email: z.string().describe("Email address of the person to share with"),
				role: z.enum(["reader", "commenter", "writer"]).default("reader").describe("Permission level: reader, commenter, or writer"),
				sendNotification: z.boolean().default(true).describe("Whether to send an email notification to the person"),
				message: z.string().optional().describe("Optional message to include in the notification email"),
			},
			async ({ fileId, email, role, sendNotification, message }) => {
				const params = new URLSearchParams({
					sendNotificationEmail: String(sendNotification),
				});
				if (message) params.set("emailMessage", message);

				const resp = await googleApiFetch(
					`https://www.googleapis.com/drive/v3/files/${fileId}/permissions?${params}`,
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							type: "user",
							role,
							emailAddress: email,
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error sharing file: ${error}` }], isError: true };
				}

				return { content: [{ type: "text", text: `Shared with ${email} as ${role}.` }] };
			},
		);

		// ===== SHEETS TOOLS =====

		this.server.tool(
			"sheets_read_range",
			"Read a range of cells from a Google Spreadsheet",
			{
				spreadsheetId: z.string().describe("The spreadsheet ID"),
				range: z.string().describe("A1 notation range, e.g. Sheet1!A1:D10"),
			},
			async ({ spreadsheetId, range }) => {
				const resp = await googleApiFetch(
					`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
					this.props!.googleAccessToken,
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error reading range: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
			},
		);

		this.server.tool(
			"sheets_write_range",
			"Write data to a range of cells in a Google Spreadsheet",
			{
				spreadsheetId: z.string().describe("The spreadsheet ID"),
				range: z.string().describe("A1 notation range, e.g. Sheet1!A1:D10"),
				values: z
					.array(z.array(z.union([z.string(), z.number(), z.boolean()])))
					.describe("2D array of values to write"),
			},
			async ({ spreadsheetId, range, values }) => {
				const resp = await googleApiFetch(
					`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
					this.props!.googleAccessToken,
					{
						method: "PUT",
						body: JSON.stringify({
							range,
							majorDimension: "ROWS",
							values,
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error writing range: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: `Updated: ${JSON.stringify(data, null, 2)}` }] };
			},
		);

		this.server.tool(
			"sheets_append_rows",
			"Append rows to a Google Spreadsheet",
			{
				spreadsheetId: z.string().describe("The spreadsheet ID"),
				range: z.string().describe("A1 notation of the table to append to, e.g. Sheet1!A:D"),
				values: z
					.array(z.array(z.union([z.string(), z.number(), z.boolean()])))
					.describe("2D array of rows to append"),
			},
			async ({ spreadsheetId, range, values }) => {
				const resp = await googleApiFetch(
					`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							range,
							majorDimension: "ROWS",
							values,
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error appending rows: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: `Appended: ${JSON.stringify(data, null, 2)}` }] };
			},
		);

		this.server.tool(
			"sheets_create_spreadsheet",
			"Create a new Google Spreadsheet",
			{
				title: z.string().describe("Spreadsheet title"),
				sheetName: z.string().default("Sheet1").describe("Name of the first sheet/tab"),
			},
			async ({ title, sheetName }) => {
				const resp = await googleApiFetch(
					"https://sheets.googleapis.com/v4/spreadsheets",
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							properties: { title },
							sheets: [{ properties: { title: sheetName } }],
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error creating spreadsheet: ${error}` }], isError: true };
				}

				const data = (await resp.json()) as { spreadsheetId: string; spreadsheetUrl: string };
				return {
					content: [
						{
							type: "text",
							text: `Spreadsheet created!\nID: ${data.spreadsheetId}\nURL: ${data.spreadsheetUrl}`,
						},
					],
				};
			},
		);

		this.server.tool(
			"sheets_add_sheet",
			"Add a new sheet/tab to an existing Google Spreadsheet",
			{
				spreadsheetId: z.string().describe("The spreadsheet ID"),
				title: z.string().describe("Name for the new sheet/tab"),
			},
			async ({ spreadsheetId, title }) => {
				const resp = await googleApiFetch(
					`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							requests: [{ addSheet: { properties: { title } } }],
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error adding sheet: ${error}` }], isError: true };
				}

				const data = await resp.json();
				return { content: [{ type: "text", text: `Sheet added: ${JSON.stringify(data, null, 2)}` }] };
			},
		);

		this.server.tool(
			"sheets_delete_sheet",
			"Delete a sheet/tab from a Google Spreadsheet",
			{
				spreadsheetId: z.string().describe("The spreadsheet ID"),
				sheetId: z.number().describe("The numeric sheet/tab ID (not the spreadsheet ID)"),
			},
			async ({ spreadsheetId, sheetId }) => {
				const resp = await googleApiFetch(
					`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							requests: [{ deleteSheet: { sheetId } }],
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error deleting sheet: ${error}` }], isError: true };
				}

				return { content: [{ type: "text", text: "Sheet deleted successfully." }] };
			},
		);

		// ===== DOCS TOOLS =====

		this.server.tool(
			"docs_read_document",
			"Read the text content of a Google Doc",
			{
				documentId: z.string().describe("The document ID"),
			},
			async ({ documentId }) => {
				const resp = await googleApiFetch(
					`https://docs.googleapis.com/v1/documents/${documentId}`,
					this.props!.googleAccessToken,
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error reading document: ${error}` }], isError: true };
				}

				const doc = (await resp.json()) as {
					title: string;
					body: {
						content: Array<{
							paragraph?: {
								elements: Array<{
									textRun?: { content: string };
								}>;
							};
						}>;
					};
				};

				// Extract text from the document body
				let text = "";
				for (const block of doc.body.content) {
					if (block.paragraph) {
						for (const element of block.paragraph.elements) {
							if (element.textRun?.content) {
								text += element.textRun.content;
							}
						}
					}
				}

				return { content: [{ type: "text", text: `Document: ${doc.title}\n\n${text}` }] };
			},
		);

		this.server.tool(
			"docs_create_document",
			"Create a new Google Doc, optionally with initial content",
			{
				title: z.string().describe("Document title"),
				initialContent: z.string().optional().describe("Optional initial text content"),
			},
			async ({ title, initialContent }) => {
				// Create the document
				const createResp = await googleApiFetch(
					"https://docs.googleapis.com/v1/documents",
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({ title }),
					},
				);

				if (!createResp.ok) {
					const error = await createResp.text();
					return { content: [{ type: "text", text: `Error creating document: ${error}` }], isError: true };
				}

				const doc = (await createResp.json()) as { documentId: string };

				// Insert initial content if provided
				if (initialContent) {
					const updateResp = await googleApiFetch(
						`https://docs.googleapis.com/v1/documents/${doc.documentId}:batchUpdate`,
						this.props!.googleAccessToken,
						{
							method: "POST",
							body: JSON.stringify({
								requests: [
									{
										insertText: {
											location: { index: 1 },
											text: initialContent,
										},
									},
								],
							}),
						},
					);

					if (!updateResp.ok) {
						const error = await updateResp.text();
						return {
							content: [
								{
									type: "text",
									text: `Document created (ID: ${doc.documentId}) but failed to insert content: ${error}`,
								},
							],
						};
					}
				}

				return {
					content: [
						{
							type: "text",
							text: `Document created!\nID: ${doc.documentId}\nURL: https://docs.google.com/document/d/${doc.documentId}/edit`,
						},
					],
				};
			},
		);

		this.server.tool(
			"docs_insert_text",
			"Insert text into a Google Doc at a specific position",
			{
				documentId: z.string().describe("The document ID"),
				text: z.string().describe("Text to insert"),
				index: z.number().default(1).describe("Position to insert at (1 = beginning of body)"),
			},
			async ({ documentId, text, index }) => {
				const resp = await googleApiFetch(
					`https://docs.googleapis.com/v1/documents/${documentId}:batchUpdate`,
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							requests: [
								{
									insertText: {
										location: { index },
										text,
									},
								},
							],
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error inserting text: ${error}` }], isError: true };
				}

				return { content: [{ type: "text", text: "Text inserted successfully." }] };
			},
		);

		this.server.tool(
			"docs_replace_text",
			"Find and replace text in a Google Doc",
			{
				documentId: z.string().describe("The document ID"),
				findText: z.string().describe("Text to find"),
				replaceText: z.string().describe("Text to replace with"),
				matchCase: z.boolean().default(true).describe("Whether to match case"),
			},
			async ({ documentId, findText, replaceText, matchCase }) => {
				const resp = await googleApiFetch(
					`https://docs.googleapis.com/v1/documents/${documentId}:batchUpdate`,
					this.props!.googleAccessToken,
					{
						method: "POST",
						body: JSON.stringify({
							requests: [
								{
									replaceAllText: {
										containsText: { text: findText, matchCase },
										replaceText,
									},
								},
							],
						}),
					},
				);

				if (!resp.ok) {
					const error = await resp.text();
					return { content: [{ type: "text", text: `Error replacing text: ${error}` }], isError: true };
				}

				const data = (await resp.json()) as {
					replies: Array<{ replaceAllText?: { occurrencesChanged: number } }>;
				};
				const changed = data.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;
				return { content: [{ type: "text", text: `Replaced ${changed} occurrence(s).` }] };
			},
		);
	}
}

export default new OAuthProvider({
	apiHandler: MyMCP.serve("/mcp"),
	apiRoute: "/mcp",
	authorizeEndpoint: "/authorize",
	clientRegistrationEndpoint: "/register",
	defaultHandler: GoogleHandler as any,
	tokenEndpoint: "/token",
});
