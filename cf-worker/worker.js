/**
 * Cloudflare Worker — proxies a GitHub Actions workflow_dispatch trigger.
 * Set the environment variable GH_TOKEN in the Worker's settings.
 * Deploy this at: workers.cloudflare.com (free tier, 100k req/day)
 */

const REPO  = 'manishsharma1999/salesforce-job-finder';
const REF   = 'main';
const WORKFLOW = 'scrape.yml';

export default {
  async fetch(request, env) {
    // Allow CORS from the GitHub Pages site
    const headers = {
      'Access-Control-Allow-Origin':  '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Content-Type': 'application/json',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers });
    }

    if (request.method !== 'POST') {
      return new Response(JSON.stringify({ error: 'POST only' }), { status: 405, headers });
    }

    const token = env.GH_TOKEN;
    if (!token) {
      return new Response(JSON.stringify({ error: 'GH_TOKEN not set' }), { status: 500, headers });
    }

    const url = `https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}/dispatches`;
    const res = await fetch(url, {
      method:  'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept':        'application/vnd.github+json',
        'Content-Type':  'application/json',
        'User-Agent':    'job-finder-worker',
      },
      body: JSON.stringify({ ref: REF }),
    });

    if (res.status === 204) {
      return new Response(JSON.stringify({ status: 'triggered' }), { status: 200, headers });
    }

    const body = await res.text();
    return new Response(JSON.stringify({ error: body }), { status: res.status, headers });
  },
};
