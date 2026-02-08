// Extend the Env interface with secrets that are set via `wrangler secret put`
declare interface Env {
	GOOGLE_CLIENT_ID: string;
	GOOGLE_CLIENT_SECRET: string;
	COOKIE_ENCRYPTION_KEY: string;
}
