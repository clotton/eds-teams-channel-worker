import { CORS_HEADERS } from "./constants";

export async function requireTurnstileHeader(request, env) {
    const token = request.headers.get('x-turnstile-token');

    if (!token) {
        return new Response(JSON.stringify({
            success: false,
            error: 'Missing x-turnstile-token header',
        }), {
            status: 400,
            headers: CORS_HEADERS(env),
        });
    }

    const valid = await verifyTurnstileToken(token, env);
    if (!valid) {
        return new Response(JSON.stringify({
            success: false,
            error: 'Invalid Turnstile token',
        }), {
            status: 403,
            headers: CORS_HEADERS(env),
        });
    }

    // Return null to indicate the request is valid
    return null;
}

async function verifyTurnstileToken(token, env) {
    if (!token) return false;

    const formData = new URLSearchParams();
    formData.append('secret', env.TURNSTILE_SECRET_KEY);
    formData.append('response', token);

    try {
        const resp = await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
            method: 'POST',
            body: formData
        });
        const data = await resp.json();
        return data.success === true;
    } catch (e) {
        console.error('Turnstile verification error:', e);
        return false;
    }
}
