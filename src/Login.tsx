import React, { useState } from 'react';

interface LoginProps {
  onLogin: (username: string) => void;
}

const Login: React.FC<LoginProps> = ({ onLogin }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  // Credenciais fixas (você pode mudar)
  const VALID_USERS = {
    'admin': 'admin123',
    'sicoob': 'cocred2026',
    'gestor': 'gestor456'
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    // Simula validação
    setTimeout(() => {
      if (VALID_USERS[username as keyof typeof VALID_USERS] === password) {
        onLogin(username);
        // Salva no localStorage para manter login
        localStorage.setItem('user', username);
        localStorage.setItem('isAuthenticated', 'true');
      } else {
        setError('Usuário ou senha inválidos!');
      }
      setLoading(false);
    }, 500);
  };

  const colors = {
    primary: '#00965E',
    primaryDark: '#006B43',
    primaryLight: '#00B37E',
    accent: '#FFCD00',
    surface: '#FFFFFF',
    background: '#F5F7FA',
    text: '#2D3748',
    textSecondary: '#718096',
    border: '#E2E8F0',
    danger: '#EF4444'
  };

  return (
    <div style={{
      minHeight: '100vh',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
      padding: '20px',
      fontFamily: "'Inter', -apple-system, sans-serif"
    }}>
      <div style={{
        background: colors.surface,
        borderRadius: '24px',
        padding: '48px',
        maxWidth: '420px',
        width: '100%',
        boxShadow: '0 24px 48px rgba(0, 0, 0, 0.2)',
        animation: 'slideUp 0.4s ease'
      }}>
        {/* Logo */}
        <div style={{ textAlign: 'center', marginBottom: '40px' }}>
          <div style={{
            width: '80px',
            height: '80px',
            background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
            borderRadius: '20px',
            margin: '0 auto 16px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 8px 24px rgba(0, 150, 94, 0.3)'
          }}>
            <svg width="40" height="40" viewBox="0 0 24 24" fill="none">
              <path d="M12 2L4 7V17L12 22L20 17V7L12 2Z" fill={colors.accent}/>
              <path d="M12 6L8 8.5V13.5L12 16L16 13.5V8.5L12 6Z" fill="white"/>
            </svg>
          </div>
          <h1 style={{
            fontSize: '24px',
            fontWeight: '700',
            color: colors.text,
            margin: 0
          }}>
            Sicoob Cocred
          </h1>
          <p style={{
            fontSize: '14px',
            color: colors.textSecondary,
            margin: '4px 0 0 0'
          }}>
            Gestão de Campanhas
          </p>
        </div>

        {/* Formulário */}
        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom: '20px' }}>
            <label style={{
              display: 'block',
              fontSize: '13px',
              fontWeight: '600',
              color: colors.text,
              marginBottom: '8px'
            }}>
              Usuário
            </label>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Digite seu usuário"
              style={{
                width: '100%',
                padding: '14px 16px',
                border: `2px solid ${error ? colors.danger : colors.border}`,
                borderRadius: '12px',
                fontSize: '14px',
                transition: 'all 0.2s',
                outline: 'none',
                backgroundColor: colors.background
              }}
              onFocus={(e) => {
                e.currentTarget.style.borderColor = colors.primary;
                e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
              }}
              onBlur={(e) => {
                e.currentTarget.style.borderColor = error ? colors.danger : colors.border;
                e.currentTarget.style.boxShadow = 'none';
              }}
              required
            />
          </div>

          <div style={{ marginBottom: '24px' }}>
            <label style={{
              display: 'block',
              fontSize: '13px',
              fontWeight: '600',
              color: colors.text,
              marginBottom: '8px'
            }}>
              Senha
            </label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Digite sua senha"
              style={{
                width: '100%',
                padding: '14px 16px',
                border: `2px solid ${error ? colors.danger : colors.border}`,
                borderRadius: '12px',
                fontSize: '14px',
                transition: 'all 0.2s',
                outline: 'none',
                backgroundColor: colors.background
              }}
              onFocus={(e) => {
                e.currentTarget.style.borderColor = colors.primary;
                e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
              }}
              onBlur={(e) => {
                e.currentTarget.style.borderColor = error ? colors.danger : colors.border;
                e.currentTarget.style.boxShadow = 'none';
              }}
              required
            />
          </div>

          {error && (
            <div style={{
              backgroundColor: '#FEE2E2',
              color: colors.danger,
              padding: '12px',
              borderRadius: '10px',
              fontSize: '13px',
              marginBottom: '20px',
              textAlign: 'center'
            }}>
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: '100%',
              padding: '16px',
              background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
              color: 'white',
              border: 'none',
              borderRadius: '12px',
              fontSize: '16px',
              fontWeight: '600',
              cursor: 'pointer',
              transition: 'all 0.2s',
              boxShadow: '0 4px 16px rgba(0, 150, 94, 0.3)',
              opacity: loading ? 0.7 : 1
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.transform = 'translateY(-2px)';
              e.currentTarget.style.boxShadow = '0 6px 20px rgba(0, 150, 94, 0.4)';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.boxShadow = '0 4px 16px rgba(0, 150, 94, 0.3)';
            }}
          >
            {loading ? 'Entrando...' : 'Entrar'}
          </button>

          {/* Credenciais de teste */}
          <div style={{
            marginTop: '24px',
            padding: '16px',
            background: colors.background,
            borderRadius: '12px',
            fontSize: '12px',
            color: colors.textSecondary,
            textAlign: 'center'
          }}>
            <p style={{ margin: '0 0 8px 0', fontWeight: '600' }}>
              Credenciais de teste:
            </p>
            <div style={{ display: 'flex', justifyContent: 'center', gap: '16px', flexWrap: 'wrap' }}>
              <span><strong>admin</strong> / admin123</span>
              <span><strong>sicoob</strong> / cocred2026</span>
            </div>
          </div>
        </form>

        <style>{`
          @keyframes slideUp {
            from {
              opacity: 0;
              transform: translateY(20px);
            }
            to {
              opacity: 1;
              transform: translateY(0);
            }
          }
        `}</style>
      </div>
    </div>
  );
};

export default Login;