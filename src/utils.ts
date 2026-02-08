/**
 * Google Workspace MCP Server - Utility Functions
 */

// Context from the auth process, encrypted & stored in the auth token
// and provided to the DurableMCP as this.props
export type Props = {
	email: string;
	name: string;
	googleAccessToken: string;
	googleRefreshToken: string;
	googleTokenExpiry: number; // Unix timestamp in seconds
};

/**
 * Constructs an authorization URL for Google OAuth.
 */
export function getUpstreamAuthorizeUrl({
	upstream_url,
	client_id,
	scope,
	redirect_uri,
	state,
}: {
	upstream_url: string;
	client_id: string;
	scope: string;
	redirect_uri: string;
	state?: string;
}) {
	const upstream = new URL(upstream_url);
	upstream.searchParams.set("client_id", client_id);
	upstream.searchParams.set("redirect_uri", redirect_uri);
	upstream.searchParams.set("scope", scope);
	upstream.searchParams.set("response_type", "code");
	upstream.searchParams.set("access_type", "offline");
	upstream.searchParams.set("prompt", "consent");
	if (state) upstream.searchParams.set("state", state);
	return upstream.href;
}

/**
 * Exchanges an authorization code for Google OAuth tokens.
 * Returns access token, refresh token, and expiry.
 */
export async function fetchUpstreamAuthToken({
	client_id,
	client_secret,
	code,
	redirect_uri,
	upstream_url,
}: {
	code: string | undefined;
	upstream_url: string;
	client_secret: string;
	redirect_uri: string;
	client_id: string;
}): Promise<
	[{ access_token: string; refresh_token: string; expires_in: number }, null] | [null, Response]
> {
	if (!code) {
		return [null, new Response("Missing code", { status: 400 })];
	}

	const resp = await fetch(upstream_url, {
		body: new URLSearchParams({
			client_id,
			client_secret,
			code,
			redirect_uri,
			grant_type: "authorization_code",
		}).toString(),
		headers: {
			"Content-Type": "application/x-www-form-urlencoded",
		},
		method: "POST",
	});

	if (!resp.ok) {
		const errorText = await resp.text();
		console.error("Google token exchange failed:", errorText);
		return [null, new Response("Failed to fetch access token", { status: 500 })];
	}

	const body = (await resp.json()) as {
		access_token: string;
		refresh_token?: string;
		expires_in: number;
	};

	if (!body.access_token) {
		return [null, new Response("Missing access token in Google response", { status: 400 })];
	}

	return [
		{
			access_token: body.access_token,
			refresh_token: body.refresh_token || "",
			expires_in: body.expires_in,
		},
		null,
	];
}

/**
 * Refreshes a Google access token using the refresh token.
 */
export async function refreshGoogleToken(
	refreshToken: string,
	clientId: string,
	clientSecret: string,
): Promise<{ access_token: string; expires_in: number; refresh_token?: string }> {
	const resp = await fetch("https://oauth2.googleapis.com/token", {
		method: "POST",
		headers: { "Content-Type": "application/x-www-form-urlencoded" },
		body: new URLSearchParams({
			grant_type: "refresh_token",
			refresh_token: refreshToken,
			client_id: clientId,
			client_secret: clientSecret,
		}).toString(),
	});

	if (!resp.ok) {
		const errorText = await resp.text();
		throw new Error(`Google token refresh failed: ${errorText}`);
	}

	return (await resp.json()) as {
		access_token: string;
		expires_in: number;
		refresh_token?: string;
	};
}

/**
 * Makes an authenticated request to a Google API.
 */
export async function googleApiFetch(
	url: string,
	accessToken: string,
	options: RequestInit = {},
): Promise<Response> {
	const response = await fetch(url, {
		...options,
		headers: {
			Authorization: `Bearer ${accessToken}`,
			"Content-Type": "application/json",
			...(options.headers || {}),
		},
	});

	return response;
}
