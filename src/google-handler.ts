import type { AuthRequest, OAuthHelpers } from "@cloudflare/workers-oauth-provider";
import { Hono } from "hono";
import { fetchUpstreamAuthToken, getUpstreamAuthorizeUrl, type Props } from "./utils";
import {
	addApprovedClient,
	bindStateToSession,
	createOAuthState,
	generateCSRFProtection,
	isClientApproved,
	OAuthError,
	renderApprovalDialog,
	validateCSRFToken,
	validateOAuthState,
} from "./workers-oauth-utils";

const GOOGLE_SCOPES = [
	"email",
	"profile",
	"https://www.googleapis.com/auth/drive",
	"https://www.googleapis.com/auth/spreadsheets",
	"https://www.googleapis.com/auth/documents",
].join(" ");

const app = new Hono<{ Bindings: Env & { OAUTH_PROVIDER: OAuthHelpers } }>();

app.get("/authorize", async (c) => {
	const oauthReqInfo = await c.env.OAUTH_PROVIDER.parseAuthRequest(c.req.raw);
	const { clientId } = oauthReqInfo;
	if (!clientId) {
		return c.text("Invalid request", 400);
	}

	// Check if client is already approved
	if (await isClientApproved(c.req.raw, clientId, c.env.COOKIE_ENCRYPTION_KEY)) {
		const { stateToken } = await createOAuthState(oauthReqInfo, c.env.OAUTH_KV);
		const { setCookie: sessionBindingCookie } = await bindStateToSession(stateToken);
		return redirectToGoogle(c.req.raw, stateToken, c.env, { "Set-Cookie": sessionBindingCookie });
	}

	// Generate CSRF protection for the approval form
	const { token: csrfToken, setCookie } = generateCSRFProtection();

	return renderApprovalDialog(c.req.raw, {
		client: await c.env.OAUTH_PROVIDER.lookupClient(clientId),
		csrfToken,
		server: {
			description:
				"Connect your Google Workspace (Drive, Sheets, Docs) to Claude via MCP.",
			logo: "https://developers.google.com/static/drive/images/drive_icon.png",
			name: "Google Workspace MCP Server",
		},
		setCookie,
		state: { oauthReqInfo },
	});
});

app.post("/authorize", async (c) => {
	try {
		const formData = await c.req.raw.formData();
		validateCSRFToken(formData, c.req.raw);

		const encodedState = formData.get("state");
		if (!encodedState || typeof encodedState !== "string") {
			return c.text("Missing state in form data", 400);
		}

		let state: { oauthReqInfo?: AuthRequest };
		try {
			state = JSON.parse(atob(encodedState));
		} catch (_e) {
			return c.text("Invalid state data", 400);
		}

		if (!state.oauthReqInfo || !state.oauthReqInfo.clientId) {
			return c.text("Invalid request", 400);
		}

		const approvedClientCookie = await addApprovedClient(
			c.req.raw,
			state.oauthReqInfo.clientId,
			c.env.COOKIE_ENCRYPTION_KEY,
		);

		const { stateToken } = await createOAuthState(state.oauthReqInfo, c.env.OAUTH_KV);
		const { setCookie: sessionBindingCookie } = await bindStateToSession(stateToken);

		const headers = new Headers();
		headers.append("Set-Cookie", approvedClientCookie);
		headers.append("Set-Cookie", sessionBindingCookie);

		return redirectToGoogle(c.req.raw, stateToken, c.env, Object.fromEntries(headers));
	} catch (error: any) {
		console.error("POST /authorize error:", error);
		if (error instanceof OAuthError) {
			return error.toResponse();
		}
		return c.text(`Internal server error: ${error.message}`, 500);
	}
});

async function redirectToGoogle(
	request: Request,
	stateToken: string,
	envBindings: Env,
	headers: Record<string, string> = {},
) {
	return new Response(null, {
		headers: {
			...headers,
			location: getUpstreamAuthorizeUrl({
				client_id: envBindings.GOOGLE_CLIENT_ID,
				redirect_uri: new URL("/callback", request.url).href,
				scope: GOOGLE_SCOPES,
				state: stateToken,
				upstream_url: "https://accounts.google.com/o/oauth2/v2/auth",
			}),
		},
		status: 302,
	});
}

/**
 * OAuth Callback from Google
 *
 * Exchanges the authorization code for tokens, fetches user info,
 * and completes the MCP OAuth flow.
 */
app.get("/callback", async (c) => {
	// Validate OAuth state with session binding
	let oauthReqInfo: AuthRequest;
	let clearSessionCookie: string;

	try {
		const result = await validateOAuthState(c.req.raw, c.env.OAUTH_KV);
		oauthReqInfo = result.oauthReqInfo;
		clearSessionCookie = result.clearCookie;
	} catch (error: any) {
		if (error instanceof OAuthError) {
			return error.toResponse();
		}
		return c.text("Internal server error", 500);
	}

	if (!oauthReqInfo.clientId) {
		return c.text("Invalid OAuth request data", 400);
	}

	// Exchange code for Google tokens
	const [tokens, errResponse] = await fetchUpstreamAuthToken({
		client_id: c.env.GOOGLE_CLIENT_ID,
		client_secret: c.env.GOOGLE_CLIENT_SECRET,
		code: c.req.query("code"),
		redirect_uri: new URL("/callback", c.req.url).href,
		upstream_url: "https://oauth2.googleapis.com/token",
	});
	if (errResponse) return errResponse;

	// Fetch user info from Google
	const userInfoResp = await fetch("https://www.googleapis.com/oauth2/v2/userinfo", {
		headers: { Authorization: `Bearer ${tokens.access_token}` },
	});

	if (!userInfoResp.ok) {
		return c.text("Failed to fetch Google user info", 500);
	}

	const userInfo = (await userInfoResp.json()) as {
		email: string;
		name: string;
		id: string;
	};

	const now = Math.floor(Date.now() / 1000);

	// Complete the OAuth flow — store Google tokens in props
	const { redirectTo } = await c.env.OAUTH_PROVIDER.completeAuthorization({
		metadata: {
			label: userInfo.name || userInfo.email,
		},
		props: {
			email: userInfo.email,
			name: userInfo.name || "",
			googleAccessToken: tokens.access_token,
			googleRefreshToken: tokens.refresh_token,
			googleTokenExpiry: now + tokens.expires_in,
		} as Props,
		request: oauthReqInfo,
		scope: oauthReqInfo.scope,
		userId: userInfo.email,
	});

	const headers = new Headers({ Location: redirectTo });
	if (clearSessionCookie) {
		headers.set("Set-Cookie", clearSessionCookie);
	}

	return new Response(null, {
		status: 302,
		headers,
	});
});

export { app as GoogleHandler };
