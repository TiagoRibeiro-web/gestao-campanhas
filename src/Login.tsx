import React, { useState } from 'react';

interface LoginProps {
  onLogin: (username: string) => void;
}

const Login: React.FC<LoginProps> = ({ onLogin }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    // Validação básica
    if (!username.trim() || !password.trim()) {
      setError('Por favor, preencha todos os campos');
      setLoading(false);
      return;
    }

    try {
      const response = await fetch('http://localhost:3002/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          username: username.trim(), 
          password: password.trim() 
        })
      });

      const result = await response.json();

      if (result.success) {
        localStorage.setItem('isAuthenticated', 'true');
        localStorage.setItem('user', username.trim());
        onLogin(username.trim());
      } else {
        setError(result.error || 'Credenciais inválidas');
      }
    } catch (error) {
      console.error('Erro no login:', error);
      setError('Erro ao conectar ao servidor. Verifique se o backend está rodando.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: '100vh',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      background: 'linear-gradient(135deg, #00965E 0%, #006B43 100%)',
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
      padding: '20px'
    }}>
      <div style={{
        background: 'white',
        borderRadius: '24px',
        padding: '48px',
        maxWidth: '420px',
        width: '100%',
        boxShadow: '0 20px 60px rgba(0, 0, 0, 0.3)',
        animation: 'slideUp 0.5s ease'
      }}>
        {/* Logo */}
        <div style={{ textAlign: 'center', marginBottom: '32px' }}>
          <div style={{
            width: '72px',
            height: '72px',
            background: '#FFCD00',
            borderRadius: '18px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            margin: '0 auto 16px auto',
            boxShadow: '0 8px 24px rgba(255, 205, 0, 0.3)'
          }}>
            <svg width="40" height="40" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M12 2L4 7V17L12 22L20 17V7L12 2Z" fill="#006B43"/>
              <path d="M12 6L8 8.5V13.5L12 16L16 13.5V8.5L12 6Z" fill="#FFCD00"/>
            </svg>
          </div>
          <h1 style={{
            fontSize: '28px',
            fontWeight: '700',
            color: '#2D3748',
            margin: 0,
            letterSpacing: '-0.5px'
          }}>
            Sicoob Cocred
          </h1>
          <p style={{
            fontSize: '14px',
            color: '#718096',
            margin: '4px 0 0 0'
          }}>
            Gestão de Campanhas & Projetos
          </p>
        </div>

        {/* Formulário */}
        <form onSubmit={handleLogin}>
          <div style={{ marginBottom: '20px' }}>
            <label style={{
              display: 'block',
              fontSize: '13px',
              fontWeight: '600',
              color: '#2D3748',
              marginBottom: '6px'
            }}>
              Usuário
            </label>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Digite seu email"
              style={{
                width: '100%',
                padding: '12px 16px',
                border: `1.5px solid ${error ? '#EF4444' : '#E2E8F0'}`,
                borderRadius: '10px',
                fontSize: '14px',
                outline: 'none',
                transition: 'all 0.2s',
                backgroundColor: '#F7FAFC'
              }}
              onFocus={(e) => {
                e.currentTarget.style.borderColor = '#00965E';
                e.currentTarget.style.boxShadow = '0 0 0 3px rgba(0, 150, 94, 0.15)';
                e.currentTarget.style.backgroundColor = 'white';
              }}
              onBlur={(e) => {
                e.currentTarget.style.borderColor = error ? '#EF4444' : '#E2E8F0';
                e.currentTarget.style.boxShadow = 'none';
                e.currentTarget.style.backgroundColor = '#F7FAFC';
              }}
              disabled={loading}
              required
            />
          </div>

          <div style={{ marginBottom: '24px' }}>
            <label style={{
              display: 'block',
              fontSize: '13px',
              fontWeight: '600',
              color: '#2D3748',
              marginBottom: '6px'
            }}>
              Senha
            </label>
            <div style={{ position: 'relative' }}>
              <input
                type={showPassword ? 'text' : 'password'}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Digite sua senha"
                style={{
                  width: '100%',
                  padding: '12px 16px',
                  paddingRight: '48px',
                  border: `1.5px solid ${error ? '#EF4444' : '#E2E8F0'}`,
                  borderRadius: '10px',
                  fontSize: '14px',
                  outline: 'none',
                  transition: 'all 0.2s',
                  backgroundColor: '#F7FAFC'
                }}
                onFocus={(e) => {
                  e.currentTarget.style.borderColor = '#00965E';
                  e.currentTarget.style.boxShadow = '0 0 0 3px rgba(0, 150, 94, 0.15)';
                  e.currentTarget.style.backgroundColor = 'white';
                }}
                onBlur={(e) => {
                  e.currentTarget.style.borderColor = error ? '#EF4444' : '#E2E8F0';
                  e.currentTarget.style.boxShadow = 'none';
                  e.currentTarget.style.backgroundColor = '#F7FAFC';
                }}
                disabled={loading}
                required
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                style={{
                  position: 'absolute',
                  right: '12px',
                  top: '50%',
                  transform: 'translateY(-50%)',
                  background: 'none',
                  border: 'none',
                  cursor: 'pointer',
                  padding: '4px',
                  color: '#718096',
                  fontSize: '18px'
                }}
              >
                {showPassword ? '👁️' : '👁️‍🗨️'}
              </button>
            </div>
          </div>

          {/* Mensagem de erro */}
          {error && (
            <div style={{
              padding: '12px 16px',
              background: '#FEE2E2',
              borderRadius: '10px',
              marginBottom: '20px',
              color: '#991B1B',
              fontSize: '13px',
              display: 'flex',
              alignItems: 'center',
              gap: '8px'
            }}>
              <span style={{ fontSize: '16px' }}>❌</span>
              {error}
            </div>
          )}

          {/* Botão de login */}
          <button
            type="submit"
            disabled={loading}
            style={{
              width: '100%',
              padding: '14px',
              background: loading 
                ? '#94A3B8' 
                : 'linear-gradient(135deg, #00965E 0%, #006B43 100%)',
              color: 'white',
              border: 'none',
              borderRadius: '10px',
              fontSize: '16px',
              fontWeight: '600',
              cursor: loading ? 'not-allowed' : 'pointer',
              transition: 'all 0.2s',
              boxShadow: loading 
                ? 'none' 
                : '0 4px 12px rgba(0, 150, 94, 0.3)',
              position: 'relative',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '8px'
            }}
            onMouseEnter={(e) => {
              if (!loading) {
                e.currentTarget.style.transform = 'translateY(-2px)';
                e.currentTarget.style.boxShadow = '0 6px 16px rgba(0, 150, 94, 0.4)';
              }
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.boxShadow = loading 
                ? 'none' 
                : '0 4px 12px rgba(0, 150, 94, 0.3)';
            }}
          >
            {loading ? (
              <>
                <span style={{
                  display: 'inline-block',
                  width: '16px',
                  height: '16px',
                  border: '2px solid white',
                  borderTopColor: 'transparent',
                  borderRadius: '50%',
                  animation: 'spin 0.8s linear infinite'
                }} />
                Entrando...
              </>
            ) : (
              'Entrar'
            )}
          </button>
        </form>

        {/* Footer */}
        <div style={{
          marginTop: '24px',
          paddingTop: '20px',
          borderTop: '1px solid #E2E8F0',
          textAlign: 'center'
        }}>
          <p style={{
            fontSize: '12px',
            color: '#718096',
            margin: 0
          }}>
            © 2026 <strong style={{ color: '#00965E' }}>Sicoob Cocred</strong> — Cooperativa de Crédito
          </p>
          <p style={{
            fontSize: '11px',
            color: '#A0AEC0',
            margin: '4px 0 0 0'
          }}>
            Sistema de Gestão de Campanhas
          </p>
        </div>
      </div>

      {/* Animações */}
      <style>{`
        @keyframes slideUp {
          from {
            opacity: 0;
            transform: translateY(20px) scale(0.98);
          }
          to {
            opacity: 1;
            transform: translateY(0) scale(1);
          }
        }

        @keyframes spin {
          to {
            transform: rotate(360deg);
          }
        }

        input:-webkit-autofill {
          -webkit-box-shadow: 0 0 0 30px #F7FAFC inset !important;
          -webkit-text-fill-color: #2D3748 !important;
        }
      `}</style>
    </div>
  );
};

export default Login;