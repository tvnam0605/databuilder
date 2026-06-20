"use client";
import { signIn } from "next-auth/react";
import { useState } from "react";

export default function LoginPage() {
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    setLoading(true);
    await signIn("google", { callbackUrl: "/tool" });
  };

  return (
    <div className="login-wrap">
      <div className="login-card">
        {/* Top accent bar */}
        <div style={{ height: 4, background: "linear-gradient(90deg, var(--or) 0%, var(--or3) 100%)" }} />

        <div className="login-top">
          {/* Logo */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 28 }}>
            <div style={{
              width: 38, height: 38, borderRadius: 10,
              background: "var(--or)", display: "flex", alignItems: "center", justifyContent: "center",
              boxShadow: "0 2px 8px rgba(217,82,4,.35)",
            }}>
              <svg width="20" height="20" fill="none" viewBox="0 0 24 24">
                <rect x="2" y="3" width="20" height="14" rx="2" stroke="white" strokeWidth="1.8"/>
                <path d="M8 21h8M12 17v4" stroke="white" strokeWidth="1.8" strokeLinecap="round"/>
                <path d="M2 8h20" stroke="white" strokeWidth="1.5" strokeOpacity=".6"/>
                <circle cx="7" cy="12" r="1.5" fill="white" fillOpacity=".7"/>
                <path d="M11 12h6" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeOpacity=".7"/>
              </svg>
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 15, letterSpacing: "-.2px" }}>Linehaul Builder</div>
              <div style={{ fontSize: 11, color: "var(--tx3)", marginTop: 1 }}>WMS TO + LH Trip Mapping</div>
            </div>
          </div>

          {/* Heading */}
          <h1 style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-.4px", marginBottom: 6 }}>
            Login to Linehaul Builder
          </h1>
          <p style={{ fontSize: 13, color: "var(--tx2)", lineHeight: 1.65, marginBottom: 24 }}>
            Use your Google Shopee / SPX Express account to access.
          </p>

          {/* Google button */}
          <button
            onClick={handleLogin}
            disabled={loading}
            style={{
              width: "100%", height: 44,
              display: "flex", alignItems: "center", justifyContent: "center", gap: 10,
              border: "1px solid var(--bd)", borderRadius: 10,
              background: loading ? "var(--srf2)" : "var(--srf)",
              cursor: loading ? "not-allowed" : "pointer",
              fontSize: 13, fontWeight: 500, color: "var(--tx)",
              transition: "all .15s",
              boxShadow: "0 1px 3px rgba(0,0,0,.08)",
            }}
            onMouseEnter={e => { if (!loading) e.currentTarget.style.background = "var(--srf2)"; }}
            onMouseLeave={e => { e.currentTarget.style.background = loading ? "var(--srf2)" : "var(--srf)"; }}
          >
            {loading ? (
              <>
                <svg style={{ animation: "spin 1s linear infinite" }} width="18" height="18" viewBox="0 0 24 24" fill="none">
                  <circle cx="12" cy="12" r="10" stroke="var(--bd2)" strokeWidth="2.5"/>
                  <path d="M12 2a10 10 0 0 1 10 10" stroke="var(--or)" strokeWidth="2.5" strokeLinecap="round"/>
                </svg>
                Logging in...
              </>
            ) : (
              <>
                <svg width="18" height="18" viewBox="0 0 48 48">
                  <path fill="#FFC107" d="M43.6 20H42v-.1H24v8h11.3c-1.6 4.7-6.1 8-11.3 8-6.6 0-12-5.4-12-12s5.4-12 12-12c3 0 5.8 1.2 8 3l5.7-5.7C34 6.1 29.3 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.7-.4-4z"/>
                  <path fill="#FF3D00" d="m6.3 14.7 6.6 4.8C14.7 15.1 19 12 24 12c3 0 5.8 1.2 8 3l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/>
                  <path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2A12 12 0 0 1 24 36c-5.2 0-9.6-3.3-11.3-8l-6.5 5C9.5 39.6 16.2 44 24 44z"/>
                  <path fill="#1976D2" d="M43.6 20H42v-.1H24v8h11.3a12 12 0 0 1-4.1 5.6l6.2 5.2C36.9 39.2 44 34 44 24c0-1.3-.1-2.7-.4-4z"/>
                </svg>
                Login with Google
              </>
            )}
          </button>

          <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
        </div>

        
      </div>
    </div>
  );
}
