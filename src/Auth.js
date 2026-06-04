import { useState } from "react";
import { supabase } from "./supabaseClient";

export default function Auth() {
  const [mode,     setMode]     = useState("login"); // login | forgot | totp_enroll | totp_verify
  const [email,    setEmail]    = useState("");
  const [password, setPassword] = useState("");
  const [totpCode, setTotpCode] = useState("");
  const [qrUrl,    setQrUrl]    = useState("");
  const [factorId, setFactorId] = useState("");
  const [loading,  setLoading]  = useState(false);
  const [message,  setMessage]  = useState(null);

  const show = (text, ok = true) => { setMessage({text, ok}); setTimeout(() => setMessage(null), 4000); };

  const handleLogin = async () => {
    if (!email || !password) return show("Preencha e-mail e senha", false);
    setLoading(true);
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) return show(error.message === "Invalid login credentials" ? "E-mail ou senha incorretos" : error.message, false);

    // Check if 2FA is required
    if (data?.session?.user) {
      const { data: factors } = await supabase.auth.mfa.listFactors();
      const totpFactor = factors?.totp?.find(f => f.status === "verified");
      if (totpFactor) {
        // Need to verify TOTP
        const { data: challenge } = await supabase.auth.mfa.challenge({ factorId: totpFactor.id });
        setFactorId(challenge?.id);
        setMode("totp_verify");
      }
      // If no 2FA factor, login proceeds normally via onAuthStateChange
    }
  };

  const handleTotpVerify = async () => {
    if (!totpCode || totpCode.length !== 6) return show("Digite o código de 6 dígitos", false);
    setLoading(true);
    const { error } = await supabase.auth.mfa.verify({
      factorId,
      challengeId: factorId,
      code: totpCode,
    });
    setLoading(false);
    if (error) return show("Código inválido ou expirado", false);
    // Auth state change will handle the rest
  };

  const handleForgot = async () => {
    if (!email) return show("Digite seu e-mail primeiro", false);
    setLoading(true);
    const { error } = await supabase.auth.resetPasswordForEmail(email);
    setLoading(false);
    if (error) show(error.message, false);
    else show("E-mail de recuperação enviado!");
  };

  // 2FA enrollment (called from inside the app after login)
  const handleEnroll = async () => {
    setLoading(true);
    const { data, error } = await supabase.auth.mfa.enroll({ factorType: "totp", issuer: "Balance App" });
    setLoading(false);
    if (error) return show(error.message, false);
    setQrUrl(data.totp.qr_code);
    setFactorId(data.id);
    setMode("totp_enroll");
  };

  const handleEnrollVerify = async () => {
    if (!totpCode || totpCode.length !== 6) return show("Digite o código de 6 dígitos", false);
    setLoading(true);
    const { error } = await supabase.auth.mfa.challengeAndVerify({ factorId, code: totpCode });
    setLoading(false);
    if (error) return show("Código inválido", false);
    show("2FA ativado com sucesso!");
    setMode("login");
  };

  return (
    <div style={s.root}>
      <style>{css}</style>
      <div style={s.card}>
        <div style={s.logo}>₿alance</div>
        <div style={s.subtitle}>Finanças Pessoais</div>

        {message && (
          <div style={{...s.msg, background: message.ok ? "#052e16" : "#1c0a0a", borderColor: message.ok ? "#166534" : "#7f1d1d"}}>
            <span style={{color: message.ok ? "#4ade80" : "#f87171"}}>{message.text}</span>
          </div>
        )}

        {/* LOGIN */}
        {mode === "login" && (
          <>
            <div style={s.field}>
              <label style={s.label}>E-mail</label>
              <input style={s.input} type="email" placeholder="seu@email.com"
                value={email} onChange={e => setEmail(e.target.value)}
                onKeyDown={e => e.key === "Enter" && handleLogin()} />
            </div>
            <div style={s.field}>
              <label style={s.label}>Senha</label>
              <input style={s.input} type="password" placeholder="••••••••"
                value={password} onChange={e => setPassword(e.target.value)}
                onKeyDown={e => e.key === "Enter" && handleLogin()} />
            </div>
            <button style={s.btn} onClick={handleLogin} disabled={loading}>
              {loading ? "Entrando..." : "Entrar"}
            </button>
            <button style={s.link} onClick={() => setMode("forgot")}>Esqueci minha senha</button>
          </>
        )}

        {/* FORGOT */}
        {mode === "forgot" && (
          <>
            <div style={s.field}>
              <label style={s.label}>E-mail</label>
              <input style={s.input} type="email" placeholder="seu@email.com"
                value={email} onChange={e => setEmail(e.target.value)} />
            </div>
            <button style={s.btn} onClick={handleForgot} disabled={loading}>
              {loading ? "Enviando..." : "Enviar link de recuperação"}
            </button>
            <button style={s.link} onClick={() => setMode("login")}>Voltar ao login</button>
          </>
        )}

        {/* TOTP VERIFY (login with 2FA) */}
        {mode === "totp_verify" && (
          <>
            <div style={s.totpInfo}>🔐 Digite o código do seu aplicativo autenticador</div>
            <div style={s.field}>
              <label style={s.label}>Código de 6 dígitos</label>
              <input style={{...s.input, textAlign:"center", fontSize:24, letterSpacing:8}}
                type="text" inputMode="numeric" maxLength={6} placeholder="000000"
                value={totpCode} onChange={e => setTotpCode(e.target.value.replace(/\D/g,""))}
                onKeyDown={e => e.key === "Enter" && handleTotpVerify()} />
            </div>
            <button style={s.btn} onClick={handleTotpVerify} disabled={loading}>
              {loading ? "Verificando..." : "Confirmar"}
            </button>
          </>
        )}

        {/* TOTP ENROLL */}
        {mode === "totp_enroll" && (
          <>
            <div style={s.totpInfo}>
              Escaneie o QR code com Google Authenticator, Authy ou similar:
            </div>
            {qrUrl && <img src={qrUrl} alt="QR Code 2FA" style={s.qr} />}
            <div style={s.field}>
              <label style={s.label}>Confirme com o código gerado</label>
              <input style={{...s.input, textAlign:"center", fontSize:24, letterSpacing:8}}
                type="text" inputMode="numeric" maxLength={6} placeholder="000000"
                value={totpCode} onChange={e => setTotpCode(e.target.value.replace(/\D/g,""))} />
            </div>
            <button style={s.btn} onClick={handleEnrollVerify} disabled={loading}>
              {loading ? "Ativando..." : "Ativar 2FA"}
            </button>
            <button style={s.link} onClick={() => setMode("login")}>Cancelar</button>
          </>
        )}

        <div style={s.privacy}>
          🔒 Dados criptografados com Row Level Security
        </div>
      </div>
    </div>
  );
}

// Exported so App.js can trigger enrollment after login
export { };

const s = {
  root:     {minHeight:"100vh",background:"#020817",display:"flex",alignItems:"center",justifyContent:"center",padding:"20px",fontFamily:"'DM Sans','Helvetica Neue',sans-serif"},
  card:     {width:"100%",maxWidth:400,background:"#04091a",border:"1px solid #0f172a",borderRadius:20,padding:"36px 28px",display:"flex",flexDirection:"column",gap:0},
  logo:     {fontSize:32,fontWeight:800,background:"linear-gradient(135deg,#60a5fa,#a78bfa)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",letterSpacing:"-0.5px",marginBottom:4},
  subtitle: {fontSize:13,color:"#475569",marginBottom:28,fontWeight:600},
  msg:      {border:"1px solid",borderRadius:10,padding:"10px 14px",fontSize:13,fontWeight:600,marginBottom:16},
  field:    {marginBottom:16},
  label:    {display:"block",fontSize:11,fontWeight:700,color:"#60a5fa",textTransform:"uppercase",letterSpacing:0.8,marginBottom:6},
  input:    {width:"100%",background:"#0a1628",border:"1.5px solid #0f172a",color:"#e2e8f0",borderRadius:10,padding:"12px 14px",fontSize:15,boxSizing:"border-box",outline:"none",fontFamily:"inherit"},
  btn:      {width:"100%",padding:"14px",borderRadius:12,border:"none",background:"linear-gradient(135deg,#3b82f6,#6366f1)",color:"#fff",fontWeight:800,fontSize:16,cursor:"pointer",marginTop:8,marginBottom:12},
  link:     {background:"none",border:"none",color:"#475569",fontSize:13,cursor:"pointer",textDecoration:"underline",padding:0,marginBottom:8},
  totpInfo: {fontSize:13,color:"#94a3b8",marginBottom:16,lineHeight:1.5,textAlign:"center"},
  qr:       {width:180,height:180,margin:"0 auto 20px",display:"block",borderRadius:12,background:"#fff",padding:8},
  privacy:  {fontSize:11,color:"#1e293b",marginTop:20,lineHeight:1.5,textAlign:"center"},
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #020817; }
  input::placeholder { color: #334155; }
  button:disabled { opacity: 0.6; cursor: not-allowed; }
`;
