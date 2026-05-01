import { useEffect, useState } from "react";
import { ACCESS_CODE_MAP } from "./accessCodes";
import { AccessSessionContext, type AccessSession } from "./accessSession";

type Props = { children: React.ReactNode };
const ACCESS_SESSION_KEY = "tax_alert_access_session";

export default function AuthGate({ children }: Props) {
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState<AccessSession | null>(null);
  const [codeInput, setCodeInput] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    try {
      const raw = localStorage.getItem(ACCESS_SESSION_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as AccessSession;
        if (parsed?.code && parsed?.salesperson && parsed?.role) {
          setSession(parsed);
        }
      }
    } finally {
      setLoading(false);
    }
  }, []);

  function loginByAccessCode() {
    const code = codeInput.trim().toUpperCase();
    if (!/^[A-Z]{13}$/.test(code)) {
      setError("请输入 13 位字母指令码");
      return;
    }
    const profile = ACCESS_CODE_MAP[code];
    if (!profile) {
      setError("指令码无效，请联系管理员确认");
      return;
    }
    const nextSession: AccessSession = { code, role: profile.role, salesperson: profile.salesperson };
    localStorage.setItem(ACCESS_SESSION_KEY, JSON.stringify(nextSession));
    setSession(nextSession);
    setError("");
  }

  function signOut() {
    localStorage.removeItem(ACCESS_SESSION_KEY);
    setSession(null);
    setCodeInput("");
    setError("");
  }

  if (loading) return <div style={{ padding: 24 }}>加载中...</div>;

  if (!session) {
    return (
      <div style={{ padding: 24, maxWidth: 420 }}>
        <h2>业务员指令码登录</h2>
        <input
          style={{ width: "100%", padding: 10, marginTop: 12 }}
          placeholder="输入13位字母指令码"
          value={codeInput}
          onChange={(e) => setCodeInput(e.target.value)}
        />
        <button style={{ marginTop: 12, padding: "8px 12px" }} onClick={loginByAccessCode}>
          登录
        </button>
        {error ? <p style={{ marginTop: 10, color: "#dc2626" }}>{error}</p> : null}
      </div>
    );
  }

  return (
    <AccessSessionContext.Provider value={session}>
      <div style={{ padding: 12, textAlign: "right" }}>
        <span style={{ marginRight: 12 }}>
          当前：{session.salesperson}（{session.role === "admin" ? "管理员" : "业务员"}）
        </span>
        <button onClick={signOut}>退出登录</button>
      </div>
      {children}
    </AccessSessionContext.Provider>
  );
}
