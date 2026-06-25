import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import Login from './Login';

interface Campanha {
  id: string;
  nome: string;
  planejadoMidia: number;
  realizadoMidia: number;
  saldoMidia: number;
  planejadoProd: number;
  realizadoProd: number;
  saldoProd: number;
  bolsa: string;
  periodo: string;
}

const dadosCompletos: Campanha[] = [
  { id: '1', nome: 'Agrishow', planejadoMidia: 132000, realizadoMidia: 131997.51, saldoMidia: 2.49, planejadoProd: 1000, realizadoProd: 4070, saldoProd: -3070, bolsa: 'Despesa', periodo: 'Q2' },
  { id: '2', nome: 'Agronegócio Copercana', planejadoMidia: 19000, realizadoMidia: 19000, saldoMidia: 0, planejadoProd: 1000, realizadoProd: 4235, saldoProd: -3235, bolsa: 'Despesa', periodo: 'Q2' },
  { id: '3', nome: 'Alco', planejadoMidia: 24000, realizadoMidia: 0, saldoMidia: 24000, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '4', nome: 'Assembleias', planejadoMidia: 15000, realizadoMidia: 15000, saldoMidia: 0, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '5', nome: 'Ativação de Natal', planejadoMidia: 52500, realizadoMidia: 0, saldoMidia: 52500, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q4' },
  { id: '6', nome: 'Brincar é cooperar', planejadoMidia: 26250, realizadoMidia: 0, saldoMidia: 26250, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q4' },
  { id: '7', nome: 'Cantata de Natal', planejadoMidia: 36750, realizadoMidia: 0, saldoMidia: 36750, planejadoProd: 6000, realizadoProd: 0, saldoProd: 6000, bolsa: 'FATES', periodo: 'Q4' },
  { id: '8', nome: 'Cãominhada', planejadoMidia: 30000, realizadoMidia: 0, saldoMidia: 30000, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q2' },
  { id: '9', nome: 'Capital de Giro', planejadoMidia: 10781.25, realizadoMidia: 0, saldoMidia: 10781.25, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Avulsa', periodo: 'Q4' },
  { id: '10', nome: 'Cinema na Praça', planejadoMidia: 210000, realizadoMidia: 68249.24, saldoMidia: 141750.76, planejadoProd: 20000, realizadoProd: 12627.82, saldoProd: 7372.18, bolsa: 'FATES', periodo: 'Q1/Q4' },
  { id: '11', nome: 'Circuito Sescoop Cultura', planejadoMidia: 21000, realizadoMidia: 0, saldoMidia: 21000, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q2/Q3/Q4' },
  { id: '12', nome: 'Clínicas Financeiras', planejadoMidia: 6540.07, realizadoMidia: 0, saldoMidia: 6540.07, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q3' },
  { id: '13', nome: 'Concurso Cultural', planejadoMidia: 2625, realizadoMidia: 0, saldoMidia: 2625, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '14', nome: 'Conexão Agro', planejadoMidia: 19500, realizadoMidia: 0, saldoMidia: 19500, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '15', nome: 'Conexão Empresa', planejadoMidia: 19500, realizadoMidia: 0, saldoMidia: 19500, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q3' },
  { id: '16', nome: 'Conta com a Cocred', planejadoMidia: 52500, realizadoMidia: 0, saldoMidia: 52500, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1/Q3' },
  { id: '17', nome: 'Coopera Saber', planejadoMidia: 26250, realizadoMidia: 0, saldoMidia: 26250, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '18', nome: 'Copa Cocred', planejadoMidia: 90562.5, realizadoMidia: 54290.3, saldoMidia: 36272.2, planejadoProd: 0, realizadoProd: 15217.62, saldoProd: -15217.62, bolsa: 'FATES', periodo: 'Q2' },
  { id: '19', nome: 'Corrida Serrana', planejadoMidia: 19500, realizadoMidia: 19500, saldoMidia: 0, planejadoProd: 2500, realizadoProd: 0, saldoProd: 2500, bolsa: 'FATES', periodo: 'Q1' },
  { id: '20', nome: 'Corrida Sertãozinho', planejadoMidia: 33381.6, realizadoMidia: 0, saldoMidia: 33381.6, planejadoProd: 2500, realizadoProd: 0, saldoProd: 2500, bolsa: 'FATES', periodo: 'Q3' },
  { id: '21', nome: 'Crédito', planejadoMidia: 811439.93, realizadoMidia: 790576.14, saldoMidia: 20863.79, planejadoProd: 124630, realizadoProd: 125434.4, saldoProd: -804.4, bolsa: 'Avulsa', periodo: 'Q1' },
  { id: '22', nome: 'Crédito Rural', planejadoMidia: 599689.93, realizadoMidia: 437735.94, saldoMidia: 161953.99, planejadoProd: 124630, realizadoProd: 120639.43, saldoProd: 3990.57, bolsa: 'Avulsa', periodo: 'Q2' },
  { id: '23', nome: 'Eventos RLZ', planejadoMidia: 1000, realizadoMidia: 0, saldoMidia: 1000, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Despesa', periodo: 'Q2' },
  { id: '24', nome: 'Expozebu', planejadoMidia: 17000, realizadoMidia: 17370.28, saldoMidia: -370.28, planejadoProd: 1000, realizadoProd: 1815, saldoProd: -815, bolsa: 'Despesa', periodo: 'Q2' },
  { id: '25', nome: 'Feira Renegociação', planejadoMidia: 8750, realizadoMidia: 0, saldoMidia: 8750, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Avulsa', periodo: 'Q3' },
  { id: '26', nome: 'FEMEC', planejadoMidia: 22000, realizadoMidia: 20185, saldoMidia: 1815, planejadoProd: 500, realizadoProd: 1815, saldoProd: -1315, bolsa: 'Despesa', periodo: 'Q1' },
  { id: '27', nome: 'Fenasucro + Carta Acordo', planejadoMidia: 94000, realizadoMidia: 0, saldoMidia: 94000, planejadoProd: 1000, realizadoProd: 0, saldoProd: 1000, bolsa: 'Despesa', periodo: 'Q3' },
  { id: '28', nome: 'Final de Ano', planejadoMidia: 676250, realizadoMidia: 0, saldoMidia: 676250, planejadoProd: 124630, realizadoProd: 0, saldoProd: 124630, bolsa: 'Avulsa', periodo: 'Q4' },
  { id: '29', nome: 'Governança (Projeto Novo)', planejadoMidia: 0, realizadoMidia: 120105.5, saldoMidia: -120105.5, planejadoProd: 0, realizadoProd: 8000, saldoProd: -8000, bolsa: 'Avulsa', periodo: 'Q1' },
  { id: '30', nome: 'Green Week', planejadoMidia: 10781.25, realizadoMidia: 0, saldoMidia: 10781.25, planejadoProd: 3000, realizadoProd: 0, saldoProd: 3000, bolsa: 'Avulsa', periodo: 'Q4' },
  { id: '31', nome: 'Inauguração Araraquara + Carta Acordo', planejadoMidia: 69804.38, realizadoMidia: 88945.28, saldoMidia: -19140.9, planejadoProd: 11000, realizadoProd: 9625, saldoProd: 1375, bolsa: 'Despesa', periodo: 'Q1' },
  { id: '32', nome: 'Indique e Ganhe', planejadoMidia: 137250, realizadoMidia: 0, saldoMidia: 137250, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Despesa', periodo: 'Q1/Q2/Q3/Q4' },
  { id: '33', nome: 'Investimento', planejadoMidia: 839689.93, realizadoMidia: 0, saldoMidia: 839689.93, planejadoProd: 124630, realizadoProd: 0, saldoProd: 124630, bolsa: 'Avulsa', periodo: 'Q3' },
  { id: '34', nome: 'Megacana', planejadoMidia: 12000, realizadoMidia: 0, saldoMidia: 12000, planejadoProd: 1000, realizadoProd: 0, saldoProd: 1000, bolsa: 'Despesa', periodo: 'Q3' },
  { id: '35', nome: 'Pedala Record', planejadoMidia: 28665, realizadoMidia: 28750, saldoMidia: -85, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Despesa', periodo: 'Q2' },
  { id: '36', nome: 'Produtos', planejadoMidia: 1344926.89, realizadoMidia: 432043.45, saldoMidia: 912883.44, planejadoProd: 150980, realizadoProd: 75914.8, saldoProd: 75065.2, bolsa: 'Avulsa', periodo: 'Q1/Q2/Q3/Q4' },
  { id: '37', nome: 'Programa Financinhas', planejadoMidia: 5512.5, realizadoMidia: 0, saldoMidia: 5512.5, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '38', nome: 'Se Liga Finanças ON', planejadoMidia: 15330, realizadoMidia: 0, saldoMidia: 15330, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'FATES', periodo: 'Q1' },
  { id: '39', nome: 'Semana Coop', planejadoMidia: 8750, realizadoMidia: 0, saldoMidia: 8750, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Avulsa', periodo: 'Q3' },
  { id: '40', nome: 'Sucessão Familiar', planejadoMidia: 26250, realizadoMidia: 0, saldoMidia: 26250, planejadoProd: 0, realizadoProd: 0, saldoProd: 0, bolsa: 'Avulsa', periodo: 'Q2 e Q3' },
  { id: '41', nome: 'Campanha Copa - Abertura de Contas', planejadoMidia: 0, realizadoMidia: 418570.47, saldoMidia: -418570.47, planejadoProd: 0, realizadoProd: 140394.63, saldoProd: -140394.63, bolsa: 'Avulsa', periodo: 'Q2' }
];

// Função para determinar o status da taxa de execução
const getStatusExecucao = (taxa: number, colors: any) => {
  if (taxa >= 90 && taxa <= 100) {
    return {
      status: 'Execução do planejado',
      cor: colors.success,
      bg: `${colors.success}15`,
      emoji: '✅'
    };
  } else if (taxa >= 70 && taxa <= 89) {
    return {
      status: 'Atenção ao planejado',
      cor: colors.warning,
      bg: `${colors.warning}15`,
      emoji: '⚠️'
    };
  } else {
    return {
      status: 'Plano não executado',
      cor: colors.danger,
      bg: `${colors.danger}15`,
      emoji: '🔴'
    };
  }
};

const App: React.FC = () => {
  // States de autenticação
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [user, setUser] = useState('');

  // States do dashboard
  const [carregando, setCarregando] = useState(false);
  const [campanhas, setCampanhas] = useState<Campanha[]>([]);
  const [filtroBolsa, setFiltroBolsa] = useState('todas');
  const [filtroPeriodo, setFiltroPeriodo] = useState('todos');
  const [busca, setBusca] = useState('');
  const [modalAberto, setModalAberto] = useState(false);
  const [editandoId, setEditandoId] = useState<string | null>(null);
  const [formData, setFormData] = useState({
    nome: '',
    planejadoMidia: 0,
    realizadoMidia: 0,
    planejadoProd: 0,
    realizadoProd: 0,
    bolsa: 'Avulsa',
    periodo: 'Q1'
  });

  // Verifica autenticação ao carregar
  useEffect(() => {
    const auth = localStorage.getItem('isAuthenticated');
    const userSaved = localStorage.getItem('user');
    if (auth === 'true' && userSaved) {
      setIsAuthenticated(true);
      setUser(userSaved);
    }
  }, []);

  // Carrega dados das campanhas
  useEffect(() => {
    const dadosSalvos = localStorage.getItem('campanhas');
    if (dadosSalvos) {
      const parsed = JSON.parse(dadosSalvos);
      setCampanhas(parsed);
    } else {
      setCampanhas(dadosCompletos);
    }
  }, []);

  useEffect(() => {
    if (campanhas.length > 0) {
      localStorage.setItem('campanhas', JSON.stringify(campanhas));
    }
  }, [campanhas]);

  // Funções de autenticação
  const handleLogin = (username: string) => {
    setUser(username);
    setIsAuthenticated(true);
  };

  const handleLogout = () => {
    localStorage.removeItem('isAuthenticated');
    localStorage.removeItem('user');
    setIsAuthenticated(false);
    setUser('');
  };

  // Função para sincronizar com SharePoint
// Função para sincronizar com SharePoint (JÁ ESTÁ CORRETA)
const sincronizarSharePoint = async () => {
  setCarregando(true);
  try {
    const response = await fetch('http://localhost:3002/api/buscar-planilha', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    const result = await response.json();
    
    if (result.success && result.data) {
      setCampanhas(result.data);
      localStorage.setItem('campanhas', JSON.stringify(result.data));
      alert(`✅ ${result.data.length} campanhas carregadas do SharePoint!`);
    } else {
      alert(`❌ Erro: ${result.error}`);
    }
  } catch (error) {
    alert('❌ Erro ao conectar com o servidor.');
  } finally {
    setCarregando(false);
  }
};

// Função para SALVAR campanhas no Excel
const salvarNoExcel = async () => {
  setCarregando(true);
  try {
    console.log('💾 Salvando campanhas no Excel...');
    
    // Preparar os dados no formato correto
    const dadosParaSalvar = campanhas.map(({ id, ...rest }) => ({
      nome: rest.nome,
      planejadoMidia: rest.planejadoMidia,
      realizadoMidia: rest.realizadoMidia,
      saldoMidia: rest.saldoMidia,
      planejadoProd: rest.planejadoProd,
      realizadoProd: rest.realizadoProd,
      saldoProd: rest.saldoProd,
      bolsa: rest.bolsa,
      periodo: rest.periodo
    }));

    // Enviar no formato esperado pelo backend
    const response = await fetch('http://localhost:3002/api/salvar-campanhas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        campanhas: dadosParaSalvar  // <-- A CHAVE CORRETA É "campanhas"
      })
    });
    
    const result = await response.json();
    
    if (result.success) {
      alert(`✅ ${result.totalRows} campanhas salvas com sucesso no Excel!`);
    } else {
      alert(`❌ Erro ao salvar: ${result.error}`);
    }
  } catch (error) {
    console.error('Erro ao salvar no Excel:', error);
    alert('❌ Erro ao conectar com o servidor. Verifique se o backend está rodando.');
  } finally {
    setCarregando(false);
  }
};

  const resetarDados = () => {
    if (window.confirm('Resetar para as 41 campanhas originais?')) {
      setCampanhas(dadosCompletos);
    }
  };

  const importarExcel = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet);

      const campanhasImportadas: Campanha[] = rows.map((row: any, index: number) => {
        const planejadoMidia = Number(row["Valor Planejado Mídia/Projeto"] || 0);
        const realizadoMidia = Number(row["Valor Realizado Mídia/Projeto"] || 0);
        const planejadoProd = Number(row["Valor Planejado Prod"] || 0);
        const realizadoProd = Number(row["Valor Realizado Prod"] || 0);
        
        return {
          id: `import_${Date.now()}_${index}`,
          nome: row["Projeto/Campanha"] || "Sem nome",
          planejadoMidia,
          realizadoMidia,
          saldoMidia: planejadoMidia - realizadoMidia,
          planejadoProd,
          realizadoProd,
          saldoProd: planejadoProd - realizadoProd,
          bolsa: row["Bolsa"] || "Avulsa",
          periodo: row["Período"] || "Q1"
        };
      });

      setCampanhas(campanhasImportadas);
    };
    reader.readAsBinaryString(file);
  };

  const exportarExcel = () => {
    const dadosExportar = campanhas.map(({ id, ...rest }) => ({
      "Projeto/Campanha": rest.nome,
      "Valor Planejado Mídia/Projeto": rest.planejadoMidia,
      "Valor Realizado Mídia/Projeto": rest.realizadoMidia,
      "Saldo Mídia/Projeto": rest.saldoMidia,
      "Valor Planejado Prod": rest.planejadoProd,
      "Valor Realizado Prod": rest.realizadoProd,
      "Saldo Prod": rest.saldoProd,
      "Bolsa": rest.bolsa,
      "Período": rest.periodo
    }));
    
    const ws = XLSX.utils.json_to_sheet(dadosExportar);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Campanhas");
    XLSX.writeFile(wb, `campanhas_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const calcularSaldos = (data: typeof formData) => ({
    ...data,
    saldoMidia: data.planejadoMidia - data.realizadoMidia,
    saldoProd: data.planejadoProd - data.realizadoProd
  });

  const adicionarCampanha = () => {
    if (!formData.nome.trim()) {
      alert('Nome é obrigatório');
      return;
    }
    const { saldoMidia, saldoProd, ...dados } = calcularSaldos(formData);
    const novaCampanha: Campanha = {
      ...dados,
      saldoMidia,
      saldoProd,
      id: Date.now().toString()
    };
    setCampanhas([...campanhas, novaCampanha]);
    fecharModal();
  };

  const editarCampanha = () => {
    if (!editandoId) return;
    const { saldoMidia, saldoProd, ...dados } = calcularSaldos(formData);
    setCampanhas(campanhas.map(camp => 
      camp.id === editandoId ? { ...dados, saldoMidia, saldoProd, id: editandoId } : camp
    ));
    fecharModal();
  };

  const excluirCampanha = (id: string) => {
    if (window.confirm('Tem certeza?')) {
      setCampanhas(campanhas.filter(camp => camp.id !== id));
    }
  };

  const abrirModalEditar = (campanha: Campanha) => {
    setEditandoId(campanha.id);
    setFormData({
      nome: campanha.nome,
      planejadoMidia: campanha.planejadoMidia,
      realizadoMidia: campanha.realizadoMidia,
      planejadoProd: campanha.planejadoProd,
      realizadoProd: campanha.realizadoProd,
      bolsa: campanha.bolsa,
      periodo: campanha.periodo
    });
    setModalAberto(true);
  };

  const abrirModalNovo = () => {
    setEditandoId(null);
    setFormData({
      nome: '',
      planejadoMidia: 0,
      realizadoMidia: 0,
      planejadoProd: 0,
      realizadoProd: 0,
      bolsa: 'Avulsa',
      periodo: 'Q1'
    });
    setModalAberto(true);
  };

  const fecharModal = () => {
    setModalAberto(false);
    setEditandoId(null);
  };

  const campanhasFiltradas = campanhas.filter(camp => {
    const matchBolsa = filtroBolsa === 'todas' || camp.bolsa === filtroBolsa;
    const matchPeriodo = filtroPeriodo === 'todos' || camp.periodo.includes(filtroPeriodo);
    const matchBusca = camp.nome.toLowerCase().includes(busca.toLowerCase());
    return matchBolsa && matchPeriodo && matchBusca;
  });

  const totais = {
    planejadoMidia: campanhasFiltradas.reduce((sum, c) => sum + c.planejadoMidia, 0),
    realizadoMidia: campanhasFiltradas.reduce((sum, c) => sum + c.realizadoMidia, 0),
    saldoMidia: campanhasFiltradas.reduce((sum, c) => sum + c.saldoMidia, 0),
    planejadoProd: campanhasFiltradas.reduce((sum, c) => sum + c.planejadoProd, 0),
    realizadoProd: campanhasFiltradas.reduce((sum, c) => sum + c.realizadoProd, 0),
    saldoProd: campanhasFiltradas.reduce((sum, c) => sum + c.saldoProd, 0)
  };

  // Paleta oficial Sicoob Cocred
  const colors = {
    primary: '#00965E',
    primaryDark: '#006B43',
    primaryLight: '#00B37E',
    accent: '#FFCD00',
    accentDark: '#E6B800',
    background: '#F5F7FA',
    surface: '#FFFFFF',
    border: '#E2E8F0',
    text: '#2D3748',
    textSecondary: '#718096',
    success: '#10B981',
    danger: '#EF4444',
    warning: '#F59E0B'
  };

  // Se não estiver autenticado, mostra a tela de login
  if (!isAuthenticated) {
    return <Login onLogin={handleLogin} />;
  }

  // Se estiver autenticado, mostra o dashboard
  return (
    <div style={{
      minHeight: '100vh',
      backgroundColor: colors.background,
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
    }}>
      {/* Header com identidade Sicoob Cocred */}
      <header style={{
        background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
        color: colors.surface,
        padding: '0',
        boxShadow: '0 4px 20px rgba(0, 150, 94, 0.15)',
        position: 'sticky',
        top: 0,
        zIndex: 100
      }}>
        <div style={{
          maxWidth: '1400px',
          margin: '0 auto',
          padding: '20px 32px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          flexWrap: 'wrap',
          gap: '16px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
            <div style={{
              width: '48px',
              height: '48px',
              background: colors.accent,
              borderRadius: '12px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              boxShadow: '0 4px 12px rgba(255, 205, 0, 0.3)'
            }}>
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M12 2L4 7V17L12 22L20 17V7L12 2Z" fill={colors.primaryDark}/>
                <path d="M12 6L8 8.5V13.5L12 16L16 13.5V8.5L12 6Z" fill={colors.accent}/>
              </svg>
            </div>
            <div>
              <h1 style={{
                fontSize: '24px',
                fontWeight: '700',
                margin: 0,
                letterSpacing: '-0.5px'
              }}>
                Sicoob Cocred
              </h1>
              <p style={{
                fontSize: '13px',
                margin: 0,
                opacity: 0.9,
                fontWeight: '500'
              }}>
                Gestão de Campanhas & Projetos
              </p>
            </div>
          </div>
          
          <div style={{
            display: 'flex',
            gap: '12px',
            alignItems: 'center'
          }}>
            <span style={{
              padding: '6px 14px',
              backgroundColor: 'rgba(255, 255, 255, 0.15)',
              borderRadius: '20px',
              fontSize: '13px',
              fontWeight: '600',
              backdropFilter: 'blur(10px)'
            }}>
              👤 {user}
            </span>
            
            <button
              onClick={handleLogout}
              style={{
                padding: '8px 16px',
                backgroundColor: 'rgba(255, 255, 255, 0.15)',
                color: 'white',
                border: 'none',
                borderRadius: '10px',
                cursor: 'pointer',
                fontSize: '13px',
                fontWeight: '600',
                transition: 'all 0.2s'
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = 'rgba(255, 255, 255, 0.25)';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = 'rgba(255, 255, 255, 0.15)';
              }}
            >
              Sair
            </button>

            <span style={{
              padding: '6px 14px',
              backgroundColor: 'rgba(255, 255, 255, 0.15)',
              borderRadius: '20px',
              fontSize: '13px',
              fontWeight: '600',
              backdropFilter: 'blur(10px)'
            }}>
              {campanhasFiltradas.length} campanhas
            </span>
          </div>
        </div>
      </header>

      <main style={{
        maxWidth: '1400px',
        margin: '0 auto',
        padding: '32px'
      }}>
        {/* Cards de Resumo com design moderno */}
        <div style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))',
          gap: '24px',
          marginBottom: '32px'
        }}>
          {/* Card Mídia */}
          <div style={{
            background: colors.surface,
            borderRadius: '16px',
            padding: '28px',
            boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04), 0 1px 2px rgba(0, 0, 0, 0.06)',
            border: `1px solid ${colors.border}`,
            transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
            position: 'relative',
            overflow: 'hidden'
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.transform = 'translateY(-4px)';
            e.currentTarget.style.boxShadow = '0 12px 24px rgba(0, 150, 94, 0.12)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.transform = 'translateY(0)';
            e.currentTarget.style.boxShadow = '0 2px 8px rgba(0, 0, 0, 0.04), 0 1px 2px rgba(0, 0, 0, 0.06)';
          }}>
            <div style={{
              position: 'absolute',
              top: 0,
              left: 0,
              right: 0,
              height: '4px',
              background: `linear-gradient(90deg, ${colors.primary} 0%, ${colors.primaryLight} 100%)`
            }} />
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '20px' }}>
              <div>
                <p style={{ fontSize: '13px', color: colors.textSecondary, fontWeight: '600', margin: 0, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Mídia / Projeto
                </p>
                <p style={{ fontSize: '28px', fontWeight: '700', color: colors.text, margin: '8px 0 0 0' }}>
                  R$ {totais.realizadoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div style={{
                width: '48px',
                height: '48px',
                borderRadius: '12px',
                background: `${colors.primary}15`,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M21 3H3C1.9 3 1 3.9 1 5V19C1 20.1 1.9 21 3 21H21C22.1 21 23 20.1 23 19V5C23 3.9 22.1 3 21 3ZM21 19H3V5H21V19ZM5 15H9V17H5V15ZM11 15H19V17H11V15ZM5 11H9V13H5V11ZM11 11H19V13H11V11ZM5 7H9V9H5V7ZM11 7H19V9H11V7Z" fill={colors.primary}/>
                </svg>
              </div>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', paddingTop: '16px', borderTop: `1px solid ${colors.border}` }}>
              <div>
                <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0' }}>Planejado</p>
                <p style={{ fontSize: '14px', fontWeight: '600', color: colors.text, margin: 0 }}>
                  R$ {totais.planejadoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div style={{ textAlign: 'right' }}>
                <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0' }}>Saldo</p>
                <p style={{
                  fontSize: '14px',
                  fontWeight: '700',
                  margin: 0,
                  color: totais.saldoMidia >= 0 ? colors.success : colors.danger
                }}>
                  {totais.saldoMidia >= 0 ? '+' : ''}R$ {totais.saldoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
            </div>
          </div>

          {/* Card Produção */}
          <div style={{
            background: colors.surface,
            borderRadius: '16px',
            padding: '28px',
            boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04), 0 1px 2px rgba(0, 0, 0, 0.06)',
            border: `1px solid ${colors.border}`,
            transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
            position: 'relative',
            overflow: 'hidden'
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.transform = 'translateY(-4px)';
            e.currentTarget.style.boxShadow = '0 12px 24px rgba(0, 150, 94, 0.12)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.transform = 'translateY(0)';
            e.currentTarget.style.boxShadow = '0 2px 8px rgba(0, 0, 0, 0.04), 0 1px 2px rgba(0, 0, 0, 0.06)';
          }}>
            <div style={{
              position: 'absolute',
              top: 0,
              left: 0,
              right: 0,
              height: '4px',
              background: `linear-gradient(90deg, ${colors.primaryLight} 0%, ${colors.accent} 100%)`
            }} />
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '20px' }}>
              <div>
                <p style={{ fontSize: '13px', color: colors.textSecondary, fontWeight: '600', margin: 0, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Produção Realizada
                </p>
                <p style={{ fontSize: '28px', fontWeight: '700', color: colors.text, margin: '8px 0 0 0' }}>
                  R$ {totais.realizadoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div style={{
                width: '48px',
                height: '48px',
                borderRadius: '12px',
                background: `${colors.accent}25`,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M20 2H4C2.9 2 2 2.9 2 4V20C2 21.1 2.9 22 4 22H20C21.1 22 22 21.1 22 20V4C22 2.9 21.1 2 20 2ZM20 20H4V4H20V20ZM6 6H18V8H6V6ZM6 10H18V12H6V10ZM6 14H18V16H6V14Z" fill={colors.accentDark}/>
                </svg>
              </div>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', paddingTop: '16px', borderTop: `1px solid ${colors.border}` }}>
              <div>
                <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0' }}>Planejado</p>
                <p style={{ fontSize: '14px', fontWeight: '600', color: colors.text, margin: 0 }}>
                  R$ {totais.planejadoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div style={{ textAlign: 'right' }}>
                <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0' }}>Saldo</p>
                <p style={{
                  fontSize: '14px',
                  fontWeight: '700',
                  margin: 0,
                  color: totais.saldoProd >= 0 ? colors.success : colors.danger
                }}>
                  {totais.saldoProd >= 0 ? '+' : ''}R$ {totais.saldoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
            </div>
          </div>

          {/* Card Total Geral */}
          <div style={{
            background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
            borderRadius: '16px',
            padding: '28px',
            boxShadow: '0 8px 24px rgba(0, 150, 94, 0.25)',
            color: colors.surface,
            position: 'relative',
            overflow: 'hidden',
            transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)'
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.transform = 'translateY(-4px)';
            e.currentTarget.style.boxShadow = '0 16px 32px rgba(0, 150, 94, 0.35)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.transform = 'translateY(0)';
            e.currentTarget.style.boxShadow = '0 8px 24px rgba(0, 150, 94, 0.25)';
          }}>
            <div style={{
              position: 'absolute',
              top: '-50%',
              right: '-20%',
              width: '200px',
              height: '200px',
              borderRadius: '50%',
              background: 'rgba(255, 205, 0, 0.1)',
              filter: 'blur(40px)'
            }} />
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '20px', position: 'relative' }}>
              <div>
                <p style={{ fontSize: '13px', opacity: 0.9, fontWeight: '600', margin: 0, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Total Geral
                </p>
                <p style={{ fontSize: '28px', fontWeight: '700', margin: '8px 0 0 0' }}>
                  R$ {(totais.realizadoMidia + totais.realizadoProd).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div style={{
                width: '48px',
                height: '48px',
                borderRadius: '12px',
                background: 'rgba(255, 255, 255, 0.2)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                backdropFilter: 'blur(10px)'
              }}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M12 2C6.48 2 2 6.48 2 12C2 17.52 6.48 22 12 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 12 2ZM13 17H11V16H9V14H11V13H13V14H15V16H13V17ZM13 12H11V7H13V12Z" fill={colors.accent}/>
                </svg>
              </div>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', paddingTop: '16px', borderTop: '1px solid rgba(255, 255, 255, 0.2)', position: 'relative' }}>
              <div>
                <p style={{ fontSize: '12px', opacity: 0.9, margin: '0 0 4px 0' }}>Planejado</p>
                <p style={{ fontSize: '14px', fontWeight: '600', margin: 0 }}>
                  R$ {(totais.planejadoMidia + totais.planejadoProd).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div style={{ textAlign: 'right' }}>
                <p style={{ fontSize: '12px', opacity: 0.9, margin: '0 0 4px 0' }}>Saldo</p>
                <p style={{
                  fontSize: '14px',
                  fontWeight: '700',
                  margin: 0,
                  color: colors.accent
                }}>
                  {(totais.saldoMidia + totais.saldoProd) >= 0 ? '+' : ''}R$ {(totais.saldoMidia + totais.saldoProd).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* GRÁFICO DE ANÁLISE POR TRIMESTRE */}
        <div style={{
          background: colors.surface,
          borderRadius: '16px',
          padding: '28px',
          marginBottom: '24px',
          boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04)',
          border: `1px solid ${colors.border}`
        }}>
          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '24px',
            flexWrap: 'wrap',
            gap: '12px'
          }}>
            <div>
              <h3 style={{
                fontSize: '18px',
                fontWeight: '700',
                color: colors.text,
                margin: 0
              }}>
                📊 Análise por Trimestre
              </h3>
              <p style={{
                fontSize: '13px',
                color: colors.textSecondary,
                margin: '4px 0 0 0'
              }}>
                Comparativo de Planejado vs Realizado por período
              </p>
            </div>
            <div style={{ display: 'flex', gap: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <div style={{ width: '12px', height: '12px', borderRadius: '3px', background: colors.primary }} />
                <span style={{ fontSize: '12px', color: colors.textSecondary }}>Planejado</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <div style={{ width: '12px', height: '12px', borderRadius: '3px', background: colors.accent }} />
                <span style={{ fontSize: '12px', color: colors.textSecondary }}>Realizado</span>
              </div>
            </div>
          </div>

          {(() => {
            const periodos = ['Q1', 'Q2', 'Q3', 'Q4'];
            const dadosPorPeriodo = periodos.map(periodo => {
              const campanhasPeriodo = campanhasFiltradas.filter(c => c.periodo.includes(periodo));
              
              const planejadoMidia = campanhasPeriodo.reduce((sum, c) => sum + c.planejadoMidia, 0);
              const realizadoMidia = campanhasPeriodo.reduce((sum, c) => sum + c.realizadoMidia, 0);
              const planejadoProd = campanhasPeriodo.reduce((sum, c) => sum + c.planejadoProd, 0);
              const realizadoProd = campanhasPeriodo.reduce((sum, c) => sum + c.realizadoProd, 0);
              
              const totalPlanejado = planejadoMidia + planejadoProd;
              const totalRealizado = realizadoMidia + realizadoProd;
              const taxaExecucao = totalPlanejado > 0 ? (totalRealizado / totalPlanejado) * 100 : 0;
              
              return {
                periodo,
                planejadoMidia,
                realizadoMidia,
                planejadoProd,
                realizadoProd,
                totalPlanejado,
                totalRealizado,
                taxaExecucao,
                totalCampanhas: campanhasPeriodo.length
              };
            });

            const maxValor = Math.max(
              ...dadosPorPeriodo.flatMap(d => [d.totalPlanejado, d.totalRealizado])
            ) * 1.2 || 1000000;

            return (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                {/* Gráfico de Barras */}
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: `repeat(${periodos.length}, 1fr)`,
                  gap: '16px',
                  alignItems: 'flex-end',
                  minHeight: '260px'
                }}>
                  {dadosPorPeriodo.map((dados, index) => {
                    const alturaPlanejado = (dados.totalPlanejado / maxValor) * 180;
                    const alturaRealizado = (dados.totalRealizado / maxValor) * 180;
                    const isEstourado = dados.totalRealizado > dados.totalPlanejado;
                    const status = getStatusExecucao(dados.taxaExecucao, colors);

                    return (
                      <div key={dados.periodo} style={{
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        height: '100%',
                        justifyContent: 'flex-end'
                      }}>
                        {/* Barras */}
                        <div style={{
                          display: 'flex',
                          alignItems: 'flex-end',
                          gap: '6px',
                          height: '200px',
                          width: '100%',
                          justifyContent: 'center',
                          position: 'relative'
                        }}>
                          {/* Barra do Planejado */}
                          <div style={{
                            width: '32px',
                            background: colors.primary,
                            height: `${Math.max(alturaPlanejado, 4)}px`,
                            borderRadius: '6px 6px 0 0',
                            transition: 'height 0.6s cubic-bezier(0.4, 0, 0.2, 1)',
                            position: 'relative',
                            boxShadow: '0 2px 8px rgba(0, 150, 94, 0.2)',
                            minHeight: '4px'
                          }}>
                            <div style={{
                              position: 'absolute',
                              top: '-20px',
                              left: '50%',
                              transform: 'translateX(-50%)',
                              fontSize: '10px',
                              fontWeight: '700',
                              color: colors.primary,
                              whiteSpace: 'nowrap'
                            }}>
                              R$ {dados.totalPlanejado.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}
                            </div>
                          </div>

                          {/* Barra do Realizado */}
                          <div style={{
                            width: '32px',
                            background: isEstourado ? colors.danger : colors.accent,
                            height: `${Math.max(alturaRealizado, 4)}px`,
                            borderRadius: '6px 6px 0 0',
                            transition: 'height 0.6s cubic-bezier(0.4, 0, 0.2, 1) 0.1s',
                            position: 'relative',
                            boxShadow: `0 2px 8px ${isEstourado ? 'rgba(239, 68, 68, 0.2)' : 'rgba(255, 205, 0, 0.3)'}`,
                            minHeight: '4px'
                          }}>
                            <div style={{
                              position: 'absolute',
                              top: '-20px',
                              left: '50%',
                              transform: 'translateX(-50%)',
                              fontSize: '10px',
                              fontWeight: '700',
                              color: isEstourado ? colors.danger : colors.accentDark,
                              whiteSpace: 'nowrap'
                            }}>
                              R$ {dados.totalRealizado.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}
                            </div>
                          </div>
                        </div>

                        {/* Label do período */}
                        <div style={{
                          marginTop: '12px',
                          textAlign: 'center',
                          width: '100%'
                        }}>
                          <div style={{
                            fontSize: '16px',
                            fontWeight: '700',
                            color: colors.text
                          }}>
                            {dados.periodo}
                          </div>
                          <div style={{
                            fontSize: '11px',
                            color: colors.textSecondary,
                            marginTop: '2px'
                          }}>
                            {dados.totalCampanhas} campanhas
                          </div>
                          <div style={{
                            fontSize: '12px',
                            fontWeight: '600',
                            color: status.cor,
                            marginTop: '4px',
                            padding: '2px 8px',
                            background: status.bg,
                            borderRadius: '12px',
                            display: 'inline-block'
                          }}>
                            {status.emoji} {dados.taxaExecucao.toFixed(1)}%
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Cards de Resumo por Trimestre */}
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: `repeat(${periodos.length}, 1fr)`,
                  gap: '12px',
                  marginTop: '8px'
                }}>
                  {dadosPorPeriodo.map((dados) => {
                    const status = getStatusExecucao(dados.taxaExecucao, colors);
                    return (
                      <div key={dados.periodo} style={{
                        background: colors.background,
                        borderRadius: '10px',
                        padding: '14px',
                        border: `1px solid ${colors.border}`
                      }}>
                        <div style={{
                          display: 'flex',
                          justifyContent: 'space-between',
                          alignItems: 'center',
                          marginBottom: '8px'
                        }}>
                          <span style={{ fontSize: '12px', fontWeight: '600', color: colors.text }}>
                            {dados.periodo}
                          </span>
                          <span style={{
                            fontSize: '11px',
                            padding: '2px 10px',
                            borderRadius: '12px',
                            background: status.bg,
                            color: status.cor,
                            fontWeight: '600'
                          }}>
                            {status.emoji} {dados.taxaExecucao.toFixed(0)}%
                          </span>
                        </div>
                        <div style={{
                          display: 'grid',
                          gridTemplateColumns: '1fr 1fr',
                          gap: '4px',
                          fontSize: '11px'
                        }}>
                          <span style={{ color: colors.textSecondary }}>📊 Planejado:</span>
                          <span style={{ fontWeight: '600', textAlign: 'right', color: colors.primary }}>
                            R$ {dados.totalPlanejado.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}
                          </span>
                          <span style={{ color: colors.textSecondary }}>✅ Realizado:</span>
                          <span style={{ 
                            fontWeight: '600', 
                            textAlign: 'right',
                            color: dados.totalRealizado > dados.totalPlanejado ? colors.danger : colors.success
                          }}>
                            R$ {dados.totalRealizado.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}
                          </span>
                          <span style={{ color: colors.textSecondary }}>💵 Saldo:</span>
                          <span style={{ 
                            fontWeight: '700', 
                            textAlign: 'right',
                            color: (dados.totalPlanejado - dados.totalRealizado) >= 0 ? colors.success : colors.danger
                          }}>
                            {((dados.totalPlanejado - dados.totalRealizado) >= 0 ? '+' : '')}
                            R$ {(dados.totalPlanejado - dados.totalRealizado).toLocaleString('pt-BR', { maximumFractionDigits: 0 })}
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })()}
        </div>

        {/* Barra de Filtros e Ações */}
        <div style={{
          background: colors.surface,
          borderRadius: '16px',
          padding: '24px',
          marginBottom: '24px',
          boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04)',
          border: `1px solid ${colors.border}`
        }}>
          <div style={{
            display: 'flex',
            flexWrap: 'wrap',
            gap: '20px',
            justifyContent: 'space-between',
            alignItems: 'flex-end'
          }}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '16px', flex: 1 }}>
              <div style={{ flex: '1 1 240px', minWidth: '200px' }}>
                <label style={{
                  display: 'block',
                  fontSize: '12px',
                  fontWeight: '600',
                  color: colors.textSecondary,
                  marginBottom: '8px',
                  textTransform: 'uppercase',
                  letterSpacing: '0.5px'
                }}>
                  Buscar Campanha
                </label>
                <div style={{ position: 'relative' }}>
                  <svg style={{
                    position: 'absolute',
                    left: '14px',
                    top: '50%',
                    transform: 'translateY(-50%)',
                    pointerEvents: 'none'
                  }} width="18" height="18" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M15.5 14H14.71L14.43 13.73C15.41 12.59 16 11.11 16 9.5C16 5.91 13.09 3 9.5 3C5.91 3 3 5.91 3 9.5C3 13.09 5.91 16 9.5 16C11.11 16 12.59 15.41 13.73 14.43L14 14.71V15.5L19 20.49L20.49 19L15.5 14ZM9.5 14C7.01 14 5 11.99 5 9.5C5 7.01 7.01 5 9.5 5C11.99 5 14 7.01 14 9.5C14 11.99 11.99 14 9.5 14Z" fill={colors.textSecondary}/>
                  </svg>
                  <input
                    type="text"
                    value={busca}
                    onChange={(e) => setBusca(e.target.value)}
                    placeholder="Digite o nome..."
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '10px 14px 10px 42px',
                      fontSize: '14px',
                      transition: 'all 0.2s ease',
                      backgroundColor: colors.background,
                      color: colors.text,
                      outline: 'none'
                    }}
                    onFocus={(e) => {
                      e.currentTarget.style.borderColor = colors.primary;
                      e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
                    }}
                    onBlur={(e) => {
                      e.currentTarget.style.borderColor = colors.border;
                      e.currentTarget.style.boxShadow = 'none';
                    }}
                  />
                </div>
              </div>
              <div style={{ flex: '0 1 180px', minWidth: '150px' }}>
                <label style={{
                  display: 'block',
                  fontSize: '12px',
                  fontWeight: '600',
                  color: colors.textSecondary,
                  marginBottom: '8px',
                  textTransform: 'uppercase',
                  letterSpacing: '0.5px'
                }}>
                  Bolsa
                </label>
                <select
                  value={filtroBolsa}
                  onChange={(e) => setFiltroBolsa(e.target.value)}
                  style={{
                    width: '100%',
                    border: `1.5px solid ${colors.border}`,
                    borderRadius: '10px',
                    padding: '10px 14px',
                    fontSize: '14px',
                    backgroundColor: colors.background,
                    color: colors.text,
                    cursor: 'pointer',
                    outline: 'none'
                  }}
                >
                  <option value="todas">Todas</option>
                  <option value="Despesa">Despesa</option>
                  <option value="FATES">FATES</option>
                  <option value="Avulsa">Avulsa</option>
                </select>
              </div>
              <div style={{ flex: '0 1 180px', minWidth: '150px' }}>
                <label style={{
                  display: 'block',
                  fontSize: '12px',
                  fontWeight: '600',
                  color: colors.textSecondary,
                  marginBottom: '8px',
                  textTransform: 'uppercase',
                  letterSpacing: '0.5px'
                }}>
                  Período
                </label>
                <select
                  value={filtroPeriodo}
                  onChange={(e) => setFiltroPeriodo(e.target.value)}
                  style={{
                    width: '100%',
                    border: `1.5px solid ${colors.border}`,
                    borderRadius: '10px',
                    padding: '10px 14px',
                    fontSize: '14px',
                    backgroundColor: colors.background,
                    color: colors.text,
                    cursor: 'pointer',
                    outline: 'none'
                  }}
                >
                  <option value="todos">Todos</option>
                  <option value="Q1">Q1</option>
                  <option value="Q2">Q2</option>
                  <option value="Q3">Q3</option>
                  <option value="Q4">Q4</option>
                </select>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
              <label style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '8px',
                padding: '10px 18px',
                backgroundColor: colors.surface,
                color: colors.primary,
                border: `1.5px solid ${colors.primary}`,
                borderRadius: '10px',
                cursor: 'pointer',
                fontWeight: '600',
                fontSize: '13px',
                transition: 'all 0.2s ease'
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = `${colors.primary}10`;
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = colors.surface;
              }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M9 16H15V10H19L12 3L5 10H9V16ZM5 18H19V20H5V18Z" fill="currentColor"/>
                </svg>
                Importar
                <input type="file" accept=".xlsx, .xls" onChange={importarExcel} style={{ display: 'none' }} />
              </label>
              <button
                onClick={sincronizarSharePoint}
                disabled={carregando}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '10px 18px',
                  backgroundColor: colors.primary,
                  color: 'white',
                  border: 'none',
                  borderRadius: '10px',
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '13px',
                  opacity: carregando ? 0.6 : 1
                }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M12 2C6.48 2 2 6.48 2 12C2 17.52 6.48 22 12 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 12 2ZM12 20C7.58 20 4 16.42 4 12C4 7.58 7.58 4 12 4C16.42 4 20 7.58 20 12C20 16.42 16.42 20 12 20ZM13 7H11V13H17V11H13V7Z" fill="currentColor"/>
                </svg>
                {carregando ? 'Sincronizando...' : 'Sync SharePoint'}
              </button>
              
              <button
                onClick={salvarNoExcel}
                disabled={carregando}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '10px 18px',
                  backgroundColor: colors.success,
                  color: 'white',
                  border: 'none',
                  borderRadius: '10px',
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '13px',
                  opacity: carregando ? 0.6 : 1,
                  boxShadow: '0 4px 12px rgba(16, 185, 129, 0.3)',
                  transition: 'all 0.2s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = 'translateY(-2px)';
                  e.currentTarget.style.boxShadow = '0 6px 16px rgba(16, 185, 129, 0.4)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 4px 12px rgba(16, 185, 129, 0.3)';
                }}
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M19 13H13V19H11V13H5V11H11V5H13V11H19V13Z" fill="currentColor"/>
                </svg>
                {carregando ? 'Salvando...' : 'Salvar no Excel'}
              </button>
              
              <button
                onClick={exportarExcel}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '10px 18px',
                  backgroundColor: colors.surface,
                  color: colors.primary,
                  border: `1.5px solid ${colors.primary}`,
                  borderRadius: '10px',
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '13px',
                  transition: 'all 0.2s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.backgroundColor = `${colors.primary}10`;
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.backgroundColor = colors.surface;
                }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M19 9H15V3H9V9H5L12 16L19 9ZM5 18V20H19V18H5Z" fill="currentColor"/>
                </svg>
                Exportar
              </button>
              
              <button
                onClick={resetarDados}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '10px 18px',
                  backgroundColor: colors.surface,
                  color: colors.danger,
                  border: `1.5px solid ${colors.danger}`,
                  borderRadius: '10px',
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '13px',
                  transition: 'all 0.2s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.backgroundColor = `${colors.danger}10`;
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.backgroundColor = colors.surface;
                }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M17.65 6.35C16.2 4.9 14.21 4 12 4C7.58 4 4.01 7.58 4.01 12C4.01 16.42 7.58 20 12 20C15.73 20 18.84 17.45 19.73 14H17.65C16.83 16.33 14.61 18 12 18C8.69 18 6 15.31 6 12C6 8.69 8.69 6 12 6C13.66 6 15.14 6.69 16.22 7.78L13 11H20V4L17.65 6.35Z" fill="currentColor"/>
                </svg>
                Resetar
              </button>
              <button
                onClick={abrirModalNovo}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '10px 20px',
                  background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
                  color: colors.surface,
                  border: 'none',
                  borderRadius: '10px',
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '13px',
                  boxShadow: '0 4px 12px rgba(0, 150, 94, 0.3)',
                  transition: 'all 0.2s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = 'translateY(-2px)';
                  e.currentTarget.style.boxShadow = '0 6px 16px rgba(0, 150, 94, 0.4)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 4px 12px rgba(0, 150, 94, 0.3)';
                }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M19 13H13V19H11V13H5V11H11V5H13V11H19V13Z" fill="currentColor"/>
                </svg>
                Nova Campanha
              </button>
            </div>
          </div>
        </div>

        {/* Tabela de Campanhas */}
        <div style={{
          background: colors.surface,
          borderRadius: '16px',
          overflow: 'hidden',
          boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04)',
          border: `1px solid ${colors.border}`
        }}>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: '1200px' }}>
              <thead>
                <tr style={{
                  background: `linear-gradient(180deg, ${colors.background} 0%, ${colors.surface} 100%)`,
                  borderBottom: `2px solid ${colors.border}`
                }}>
                  {['Campanha', 'Planejado Mídia', 'Realizado Mídia', 'Saldo Mídia', 'Planejado Prod', 'Realizado Prod', 'Saldo Prod', 'Bolsa', 'Período', 'Ações'].map((header, i) => (
                    <th key={i} style={{
                      padding: '16px 20px',
                      textAlign: i === 0 || i === 7 || i === 8 || i === 9 ? 'left' : 'right',
                      fontSize: '11px',
                      fontWeight: '700',
                      color: colors.textSecondary,
                      textTransform: 'uppercase',
                      letterSpacing: '0.8px'
                    }}>
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {campanhasFiltradas.map((camp, index) => (
                  <tr
                    key={camp.id}
                    style={{
                      borderBottom: `1px solid ${colors.border}`,
                      transition: 'background-color 0.15s ease',
                      backgroundColor: index % 2 === 0 ? colors.surface : colors.background
                    }}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.backgroundColor = `${colors.primary}08`;
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.backgroundColor = index % 2 === 0 ? colors.surface : colors.background;
                    }}
                  >
                    <td style={{
                      padding: '16px 20px',
                      fontSize: '14px',
                      fontWeight: '600',
                      color: colors.text
                    }}>
                      {camp.nome}
                    </td>
                    <td style={{ padding: '16px 20px', fontSize: '13px', textAlign: 'right', color: colors.text, fontVariantNumeric: 'tabular-nums' }}>
                      R$ {camp.planejadoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </td>
                    <td style={{ padding: '16px 20px', fontSize: '13px', textAlign: 'right', color: colors.text, fontVariantNumeric: 'tabular-nums' }}>
                      R$ {camp.realizadoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </td>
                    <td style={{
                      padding: '16px 20px',
                      fontSize: '13px',
                      textAlign: 'right',
                      fontWeight: '700',
                      color: camp.saldoMidia >= 0 ? colors.success : colors.danger,
                      fontVariantNumeric: 'tabular-nums'
                    }}>
                      {camp.saldoMidia >= 0 ? '+' : ''}R$ {camp.saldoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </td>
                    <td style={{ padding: '16px 20px', fontSize: '13px', textAlign: 'right', color: colors.text, fontVariantNumeric: 'tabular-nums' }}>
                      R$ {camp.planejadoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </td>
                    <td style={{ padding: '16px 20px', fontSize: '13px', textAlign: 'right', color: colors.text, fontVariantNumeric: 'tabular-nums' }}>
                      R$ {camp.realizadoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </td>
                    <td style={{
                      padding: '16px 20px',
                      fontSize: '13px',
                      textAlign: 'right',
                      fontWeight: '700',
                      color: camp.saldoProd >= 0 ? colors.success : colors.danger,
                      fontVariantNumeric: 'tabular-nums'
                    }}>
                      {camp.saldoProd >= 0 ? '+' : ''}R$ {camp.saldoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </td>
                    <td style={{ padding: '16px 20px' }}>
                      <span style={{
                        padding: '5px 12px',
                        borderRadius: '20px',
                        fontSize: '11px',
                        fontWeight: '700',
                        textTransform: 'uppercase',
                        letterSpacing: '0.5px',
                        backgroundColor: camp.bolsa === 'Despesa' ? '#FEE2E2' : camp.bolsa === 'FATES' ? '#DBEAFE' : '#FEF3C7',
                        color: camp.bolsa === 'Despesa' ? '#991B1B' : camp.bolsa === 'FATES' ? '#1E40AF' : '#92400E'
                      }}>
                        {camp.bolsa}
                      </span>
                    </td>
                    <td style={{ padding: '16px 20px' }}>
                      <span style={{
                        padding: '5px 12px',
                        borderRadius: '20px',
                        fontSize: '11px',
                        fontWeight: '700',
                        backgroundColor: `${colors.primary}15`,
                        color: colors.primary
                      }}>
                        {camp.periodo}
                      </span>
                    </td>
                    <td style={{ padding: '16px 20px' }}>
                      <div style={{ display: 'flex', gap: '4px' }}>
                        <button
                          onClick={() => abrirModalEditar(camp)}
                          style={{
                            background: 'none',
                            border: 'none',
                            padding: '6px',
                            borderRadius: '6px',
                            cursor: 'pointer',
                            color: colors.primary,
                            transition: 'background-color 0.2s'
                          }}
                          onMouseEnter={(e) => e.currentTarget.style.backgroundColor = `${colors.primary}15`}
                          onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                          title="Editar"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path d="M3 17.25V21H6.75L17.81 9.94L14.06 6.19L3 17.25ZM20.71 7.04C21.1 6.65 21.1 6.02 20.71 5.63L18.37 3.29C17.98 2.9 17.35 2.9 16.96 3.29L15.13 5.12L18.88 8.87L20.71 7.04Z" fill="currentColor"/>
                          </svg>
                        </button>
                        <button
                          onClick={() => excluirCampanha(camp.id)}
                          style={{
                            background: 'none',
                            border: 'none',
                            padding: '6px',
                            borderRadius: '6px',
                            cursor: 'pointer',
                            color: colors.danger,
                            transition: 'background-color 0.2s'
                          }}
                          onMouseEnter={(e) => e.currentTarget.style.backgroundColor = `${colors.danger}15`}
                          onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                          title="Excluir"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path d="M6 19C6 20.1 6.9 21 8 21H16C17.1 21 18 20.1 18 19V7H6V19ZM19 4H15.5L14.5 3H9.5L8.5 4H5V6H19V4Z" fill="currentColor"/>
                          </svg>
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {campanhasFiltradas.length === 0 && (
            <div style={{
              textAlign: 'center',
              padding: '80px 20px',
              color: colors.textSecondary
            }}>
              <svg width="64" height="64" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ opacity: 0.3, marginBottom: '16px' }}>
                <path d="M20 6H12L10 4H4C2.9 4 2 4.9 2 6V18C2 19.1 2.9 20 4 20H20C21.1 20 22 19.1 22 18V8C22 6.9 21.1 6 20 6ZM20 18H4V8H20V18Z" fill="currentColor"/>
              </svg>
              <p style={{ fontSize: '16px', fontWeight: '500', margin: '0 0 8px 0' }}>Nenhuma campanha encontrada</p>
              <p style={{ fontSize: '14px', margin: 0, opacity: 0.7 }}>Tente ajustar os filtros ou criar uma nova campanha</p>
            </div>
          )}
        </div>

        {/* Footer com branding e explicação da Taxa de Execução */}
        <footer style={{
          marginTop: '40px',
          padding: '24px',
          background: colors.surface,
          borderRadius: '16px',
          border: `1px solid ${colors.border}`,
          textAlign: 'center',
          color: colors.textSecondary,
          fontSize: '13px'
        }}>
          {/* Totalizador rápido */}
          <div style={{
            display: 'flex',
            justifyContent: 'center',
            gap: '32px',
            marginBottom: '24px',
            flexWrap: 'wrap',
            paddingBottom: '20px',
            borderBottom: `1px solid ${colors.border}`
          }}>
            <div style={{ textAlign: 'center' }}>
              <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0', fontWeight: '600', textTransform: 'uppercase' }}>
                Total Executado
              </p>
              <p style={{ fontSize: '20px', fontWeight: '700', color: colors.primary, margin: 0 }}>
                R$ {(totais.realizadoMidia + totais.realizadoProd).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
              </p>
            </div>
            <div style={{ textAlign: 'center' }}>
              <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0', fontWeight: '600', textTransform: 'uppercase' }}>
                Taxa de Execução
              </p>
              {(() => {
                const taxa = ((totais.realizadoMidia + totais.realizadoProd) / ((totais.planejadoMidia + totais.planejadoProd) || 1) * 100);
                const status = getStatusExecucao(taxa, colors);
                return (
                  <p style={{ 
                    fontSize: '20px', 
                    fontWeight: '700', 
                    color: status.cor,
                    margin: 0
                  }}>
                    {taxa.toFixed(1)}%
                  </p>
                );
              })()}
            </div>
            <div style={{ textAlign: 'center' }}>
              <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0', fontWeight: '600', textTransform: 'uppercase' }}>
                Status do Orçamento
              </p>
              {(() => {
                const taxa = ((totais.realizadoMidia + totais.realizadoProd) / ((totais.planejadoMidia + totais.planejadoProd) || 1) * 100);
                const status = getStatusExecucao(taxa, colors);
                return (
                  <p style={{ 
                    fontSize: '16px', 
                    fontWeight: '700', 
                    color: status.cor,
                    margin: 0,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '8px'
                  }}>
                    <span>{status.emoji}</span>
                    {status.status}
                  </p>
                );
              })()}
            </div>
            <div style={{ textAlign: 'center' }}>
              <p style={{ fontSize: '12px', color: colors.textSecondary, margin: '0 0 4px 0', fontWeight: '600', textTransform: 'uppercase' }}>
                Saldo Restante
              </p>
              <p style={{ fontSize: '20px', fontWeight: '700', color: colors.primary, margin: 0 }}>
                R$ {((totais.planejadoMidia + totais.planejadoProd) - (totais.realizadoMidia + totais.realizadoProd)).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
              </p>
            </div>
          </div>

          {/* Barra de progresso da Taxa de Execução */}
          <div style={{
            background: colors.surface,
            borderRadius: '10px',
            padding: '16px 20px',
            marginBottom: '16px',
            border: `1px solid ${colors.border}`
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px' }}>
              <span style={{ fontSize: '13px', fontWeight: '600', color: colors.text }}>
                Taxa de Execução
              </span>
              {(() => {
                const taxa = ((totais.realizadoMidia + totais.realizadoProd) / ((totais.planejadoMidia + totais.planejadoProd) || 1) * 100);
                const status = getStatusExecucao(taxa, colors);
                return (
                  <span style={{ 
                    fontSize: '14px', 
                    fontWeight: '700', 
                    color: status.cor
                  }}>
                    {taxa.toFixed(1)}%
                  </span>
                );
              })()}
            </div>
            <div style={{
              width: '100%',
              height: '10px',
              background: colors.border,
              borderRadius: '5px',
              overflow: 'hidden',
              position: 'relative'
            }}>
              {(() => {
                const taxa = ((totais.realizadoMidia + totais.realizadoProd) / ((totais.planejadoMidia + totais.planejadoProd) || 1) * 100);
                const status = getStatusExecucao(taxa, colors);
                return (
                  <div style={{
                    width: `${Math.min(taxa, 100)}%`,
                    height: '100%',
                    background: status.cor,
                    borderRadius: '5px',
                    transition: 'width 0.6s cubic-bezier(0.4, 0, 0.2, 1)'
                  }} />
                );
              })()}
            </div>
            <div style={{
              display: 'flex',
              justifyContent: 'space-between',
              marginTop: '6px',
              fontSize: '11px',
              color: colors.textSecondary
            }}>
              <span>🔴 0%</span>
              <span>🟡 70%</span>
              <span>🟢 90%</span>
              <span>100%</span>
            </div>
          </div>

          {/* Explicação da Taxa de Execução */}
          <div style={{
            background: colors.background,
            borderRadius: '12px',
            padding: '20px',
            marginBottom: '20px',
            textAlign: 'left'
          }}>
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              marginBottom: '16px'
            }}>
              <span style={{ fontSize: '20px' }}>📖</span>
              <h4 style={{
                fontSize: '14px',
                fontWeight: '700',
                color: colors.text,
                margin: 0
              }}>
                O que significa a Taxa de Execução?
              </h4>
            </div>
            <p style={{
              fontSize: '13px',
              color: colors.textSecondary,
              margin: '0 0 16px 0',
              lineHeight: '1.5'
            }}>
              A <strong style={{ color: colors.primary }}>Taxa de Execução</strong> mostra quanto do orçamento planejado foi realmente executado. 
              Calculada como: <strong>(Valor Realizado ÷ Valor Planejado) × 100</strong>
            </p>
            
            {/* Legenda com novos critérios */}
            <div style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
              gap: '12px',
              marginBottom: '16px'
            }}>
              <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                padding: '8px 12px',
                background: `${colors.success}10`,
                borderRadius: '8px',
                border: `1px solid ${colors.success}30`
              }}>
                <span style={{ fontSize: '20px' }}>🟢</span>
                <div>
                  <p style={{ fontSize: '12px', fontWeight: '700', color: colors.success, margin: 0 }}>90% - 100%</p>
                  <p style={{ fontSize: '11px', color: colors.textSecondary, margin: 0 }}>Execução do planejado</p>
                </div>
              </div>
              <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                padding: '8px 12px',
                background: `${colors.warning}10`,
                borderRadius: '8px',
                border: `1px solid ${colors.warning}30`
              }}>
                <span style={{ fontSize: '20px' }}>🟡</span>
                <div>
                  <p style={{ fontSize: '12px', fontWeight: '700', color: colors.warning, margin: 0 }}>70% - 89%</p>
                  <p style={{ fontSize: '11px', color: colors.textSecondary, margin: 0 }}>Atenção ao planejado</p>
                </div>
              </div>
              <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                padding: '8px 12px',
                background: `${colors.danger}10`,
                borderRadius: '8px',
                border: `1px solid ${colors.danger}30`
              }}>
                <span style={{ fontSize: '20px' }}>🔴</span>
                <div>
                  <p style={{ fontSize: '12px', fontWeight: '700', color: colors.danger, margin: 0 }}>Abaixo de 70%</p>
                  <p style={{ fontSize: '11px', color: colors.textSecondary, margin: 0 }}>Plano não executado</p>
                </div>
              </div>
            </div>

            {/* Exemplo prático */}
            <div style={{
              padding: '12px',
              background: `${colors.primary}08`,
              borderRadius: '8px',
              border: `1px solid ${colors.border}`
            }}>
              <p style={{ fontSize: '12px', color: colors.textSecondary, margin: 0, lineHeight: '1.5' }}>
                <strong style={{ color: colors.primary }}>💡 Exemplo prático:</strong> 
                {(() => {
                  const taxa = 75;
                  const status = getStatusExecucao(taxa, colors);
                  return (
                    <span>
                      {' '}Se sua campanha tem taxa de <strong>{taxa}%</strong>, ela está em{' '}
                      <strong style={{ color: status.cor }}>{status.emoji} {status.status}</strong>
                      {' '}(faixa amarela - entre 70% e 89%).
                    </span>
                  );
                })()}
              </p>
            </div>
          </div>

          {/* Branding */}
          <p style={{ margin: 0 }}>
            © 2026 <strong style={{ color: colors.primary }}>Sicoob Cocred</strong> — Cooperativa de Crédito
          </p>
          <p style={{ margin: '4px 0 0 0', fontSize: '12px', opacity: 0.7 }}>
            Sistema de Gestão de Campanhas | Dados atualizados em tempo real
          </p>
        </footer>
      </main>

      {/* Modal com Pré-visualização dos Saldos */}
      {modalAberto && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000,
            backdropFilter: 'blur(4px)',
            animation: 'fadeIn 0.2s ease'
          }}
          onClick={fecharModal}
        >
          <div
            style={{
              background: colors.surface,
              borderRadius: '20px',
              padding: '32px',
              maxWidth: '700px',
              width: '90%',
              maxHeight: '90vh',
              overflow: 'auto',
              boxShadow: '0 24px 48px rgba(0, 0, 0, 0.2)',
              animation: 'slideUp 0.3s ease'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              marginBottom: '28px',
              paddingBottom: '20px',
              borderBottom: `1px solid ${colors.border}`
            }}>
              <div>
                <h2 style={{
                  fontSize: '22px',
                  fontWeight: '700',
                  margin: 0,
                  color: colors.text
                }}>
                  {editandoId ? 'Editar Campanha' : 'Nova Campanha'}
                </h2>
                <p style={{
                  fontSize: '13px',
                  color: colors.textSecondary,
                  margin: '4px 0 0 0'
                }}>
                  {editandoId ? 'Atualize os dados da campanha' : 'Preencha os dados da nova campanha'}
                </p>
              </div>
              <button
                onClick={fecharModal}
                style={{
                  background: 'none',
                  border: 'none',
                  padding: '8px',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  color: colors.textSecondary,
                  transition: 'all 0.2s'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.backgroundColor = colors.background;
                  e.currentTarget.style.color = colors.text;
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.backgroundColor = 'transparent';
                  e.currentTarget.style.color = colors.textSecondary;
                }}
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M19 6.41L17.59 5L12 10.59L6.41 5L5 6.41L10.59 12L5 17.59L6.41 19L12 13.41L17.59 19L19 17.59L13.41 12L19 6.41Z" fill="currentColor"/>
                </svg>
              </button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
              <div>
                <label style={{
                  display: 'block',
                  fontSize: '13px',
                  fontWeight: '600',
                  marginBottom: '8px',
                  color: colors.text
                }}>
                  Nome da Campanha *
                </label>
                <input
                  type="text"
                  value={formData.nome}
                  onChange={(e) => setFormData({ ...formData, nome: e.target.value })}
                  style={{
                    width: '100%',
                    border: `1.5px solid ${colors.border}`,
                    borderRadius: '10px',
                    padding: '12px 14px',
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
                    e.currentTarget.style.borderColor = colors.border;
                    e.currentTarget.style.boxShadow = 'none';
                  }}
                  placeholder="Ex: Campanha de Natal"
                />
              </div>

              <div style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: '16px',
                padding: '20px',
                background: colors.background,
                borderRadius: '12px'
              }}>
                <div>
                  <label style={{
                    display: 'block',
                    fontSize: '13px',
                    fontWeight: '600',
                    marginBottom: '8px',
                    color: colors.text
                  }}>
                    Planejado Mídia
                  </label>
                  <input
                    type="number"
                    value={formData.planejadoMidia}
                    onChange={(e) => setFormData({ ...formData, planejadoMidia: parseFloat(e.target.value) || 0 })}
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '12px 14px',
                      fontSize: '14px',
                      outline: 'none',
                      backgroundColor: colors.surface
                    }}
                    onFocus={(e) => {
                      e.currentTarget.style.borderColor = colors.primary;
                      e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
                    }}
                    onBlur={(e) => {
                      e.currentTarget.style.borderColor = colors.border;
                      e.currentTarget.style.boxShadow = 'none';
                    }}
                  />
                </div>
                <div>
                  <label style={{
                    display: 'block',
                    fontSize: '13px',
                    fontWeight: '600',
                    marginBottom: '8px',
                    color: colors.text
                  }}>
                    Realizado Mídia
                  </label>
                  <input
                    type="number"
                    value={formData.realizadoMidia}
                    onChange={(e) => setFormData({ ...formData, realizadoMidia: parseFloat(e.target.value) || 0 })}
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '12px 14px',
                      fontSize: '14px',
                      outline: 'none',
                      backgroundColor: colors.surface
                    }}
                    onFocus={(e) => {
                      e.currentTarget.style.borderColor = colors.primary;
                      e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
                    }}
                    onBlur={(e) => {
                      e.currentTarget.style.borderColor = colors.border;
                      e.currentTarget.style.boxShadow = 'none';
                    }}
                  />
                </div>
              </div>

              <div style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: '16px',
                padding: '20px',
                background: colors.background,
                borderRadius: '12px'
              }}>
                <div>
                  <label style={{
                    display: 'block',
                    fontSize: '13px',
                    fontWeight: '600',
                    marginBottom: '8px',
                    color: colors.text
                  }}>
                    Planejado Produção
                  </label>
                  <input
                    type="number"
                    value={formData.planejadoProd}
                    onChange={(e) => setFormData({ ...formData, planejadoProd: parseFloat(e.target.value) || 0 })}
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '12px 14px',
                      fontSize: '14px',
                      outline: 'none',
                      backgroundColor: colors.surface
                    }}
                    onFocus={(e) => {
                      e.currentTarget.style.borderColor = colors.primary;
                      e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
                    }}
                    onBlur={(e) => {
                      e.currentTarget.style.borderColor = colors.border;
                      e.currentTarget.style.boxShadow = 'none';
                    }}
                  />
                </div>
                <div>
                  <label style={{
                    display: 'block',
                    fontSize: '13px',
                    fontWeight: '600',
                    marginBottom: '8px',
                    color: colors.text
                  }}>
                    Realizado Produção
                  </label>
                  <input
                    type="number"
                    value={formData.realizadoProd}
                    onChange={(e) => setFormData({ ...formData, realizadoProd: parseFloat(e.target.value) || 0 })}
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '12px 14px',
                      fontSize: '14px',
                      outline: 'none',
                      backgroundColor: colors.surface
                    }}
                    onFocus={(e) => {
                      e.currentTarget.style.borderColor = colors.primary;
                      e.currentTarget.style.boxShadow = `0 0 0 3px ${colors.primary}20`;
                    }}
                    onBlur={(e) => {
                      e.currentTarget.style.borderColor = colors.border;
                      e.currentTarget.style.boxShadow = 'none';
                    }}
                  />
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                <div>
                  <label style={{
                    display: 'block',
                    fontSize: '13px',
                    fontWeight: '600',
                    marginBottom: '8px',
                    color: colors.text
                  }}>
                    Bolsa
                  </label>
                  <select
                    value={formData.bolsa}
                    onChange={(e) => setFormData({ ...formData, bolsa: e.target.value })}
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '12px 14px',
                      fontSize: '14px',
                      backgroundColor: colors.background,
                      color: colors.text,
                      cursor: 'pointer',
                      outline: 'none'
                    }}
                  >
                    <option value="Despesa">Despesa</option>
                    <option value="FATES">FATES</option>
                    <option value="Avulsa">Avulsa</option>
                  </select>
                </div>
                <div>
                  <label style={{
                    display: 'block',
                    fontSize: '13px',
                    fontWeight: '600',
                    marginBottom: '8px',
                    color: colors.text
                  }}>
                    Período
                  </label>
                  <select
                    value={formData.periodo}
                    onChange={(e) => setFormData({ ...formData, periodo: e.target.value })}
                    style={{
                      width: '100%',
                      border: `1.5px solid ${colors.border}`,
                      borderRadius: '10px',
                      padding: '12px 14px',
                      fontSize: '14px',
                      backgroundColor: colors.background,
                      color: colors.text,
                      cursor: 'pointer',
                      outline: 'none'
                    }}
                  >
                    <option value="Q1">Q1</option>
                    <option value="Q2">Q2</option>
                    <option value="Q3">Q3</option>
                    <option value="Q4">Q4</option>
                    <option value="Q1/Q2">Q1/Q2</option>
                    <option value="Q1/Q4">Q1/Q4</option>
                    <option value="Q2/Q3/Q4">Q2/Q3/Q4</option>
                    <option value="Q1/Q2/Q3/Q4">Todos trimestres</option>
                  </select>
                </div>
              </div>

              {/* PRÉ-VISUALIZAÇÃO DOS SALDOS */}
              <div style={{
                background: `linear-gradient(135deg, ${colors.background} 0%, ${colors.surface} 100%)`,
                borderRadius: '16px',
                padding: '20px',
                border: `2px solid ${colors.primary}25`,
                marginTop: '8px'
              }}>
                <div style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                  marginBottom: '16px'
                }}>
                  <span style={{ fontSize: '18px' }}>📊</span>
                  <h3 style={{
                    fontSize: '14px',
                    fontWeight: '700',
                    color: colors.text,
                    margin: 0
                  }}>
                    Pré-visualização dos Saldos
                  </h3>
                  <span style={{
                    marginLeft: 'auto',
                    fontSize: '11px',
                    padding: '4px 12px',
                    background: colors.primary,
                    color: 'white',
                    borderRadius: '12px',
                    fontWeight: '600'
                  }}>
                    Atualizado em tempo real
                  </span>
                </div>

                {(() => {
                  const baseCampanhas = editandoId 
                    ? campanhasFiltradas.filter(c => c.id !== editandoId)
                    : campanhasFiltradas;

                  const totaisBase = {
                    planejadoMidia: baseCampanhas.reduce((sum, c) => sum + c.planejadoMidia, 0),
                    realizadoMidia: baseCampanhas.reduce((sum, c) => sum + c.realizadoMidia, 0),
                    planejadoProd: baseCampanhas.reduce((sum, c) => sum + c.planejadoProd, 0),
                    realizadoProd: baseCampanhas.reduce((sum, c) => sum + c.realizadoProd, 0)
                  };

                  const totaisComNova = {
                    planejadoMidia: totaisBase.planejadoMidia + (formData.planejadoMidia || 0),
                    realizadoMidia: totaisBase.realizadoMidia + (formData.realizadoMidia || 0),
                    saldoMidia: (totaisBase.planejadoMidia + (formData.planejadoMidia || 0)) - 
                                (totaisBase.realizadoMidia + (formData.realizadoMidia || 0)),
                    planejadoProd: totaisBase.planejadoProd + (formData.planejadoProd || 0),
                    realizadoProd: totaisBase.realizadoProd + (formData.realizadoProd || 0),
                    saldoProd: (totaisBase.planejadoProd + (formData.planejadoProd || 0)) - 
                               (totaisBase.realizadoProd + (formData.realizadoProd || 0))
                  };

                  const totalGeralComNova = totaisComNova.realizadoMidia + totaisComNova.realizadoProd;
                  const totalPlanejadoComNova = totaisComNova.planejadoMidia + totaisComNova.planejadoProd;
                  const taxaExecucao = totalPlanejadoComNova > 0 
                    ? (totalGeralComNova / totalPlanejadoComNova * 100) 
                    : 0;

                  const status = getStatusExecucao(taxaExecucao, colors);

                  return (
                    <>
                      <div style={{
                        display: 'grid',
                        gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))',
                        gap: '12px',
                        marginBottom: '16px'
                      }}>
                        <div style={{
                          background: colors.surface,
                          borderRadius: '10px',
                          padding: '12px',
                          border: `1px solid ${colors.border}`
                        }}>
                          <p style={{ fontSize: '11px', color: colors.textSecondary, margin: '0 0 4px 0', fontWeight: '600' }}>
                            📺 Mídia
                          </p>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                            <span style={{ color: colors.textSecondary }}>Realizado:</span>
                            <strong style={{ color: colors.primary }}>
                              R$ {totaisComNova.realizadoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </strong>
                          </div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', marginTop: '2px' }}>
                            <span style={{ color: colors.textSecondary }}>Saldo:</span>
                            <strong style={{ 
                              color: totaisComNova.saldoMidia >= 0 ? colors.success : colors.danger 
                            }}>
                              {totaisComNova.saldoMidia >= 0 ? '+' : ''}
                              R$ {totaisComNova.saldoMidia.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </strong>
                          </div>
                        </div>

                        <div style={{
                          background: colors.surface,
                          borderRadius: '10px',
                          padding: '12px',
                          border: `1px solid ${colors.border}`
                        }}>
                          <p style={{ fontSize: '11px', color: colors.textSecondary, margin: '0 0 4px 0', fontWeight: '600' }}>
                            🏭 Produção
                          </p>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                            <span style={{ color: colors.textSecondary }}>Realizado:</span>
                            <strong style={{ color: colors.primary }}>
                              R$ {totaisComNova.realizadoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </strong>
                          </div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', marginTop: '2px' }}>
                            <span style={{ color: colors.textSecondary }}>Saldo:</span>
                            <strong style={{ 
                              color: totaisComNova.saldoProd >= 0 ? colors.success : colors.danger 
                            }}>
                              {totaisComNova.saldoProd >= 0 ? '+' : ''}
                              R$ {totaisComNova.saldoProd.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </strong>
                          </div>
                        </div>

                        <div style={{
                          background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
                          borderRadius: '10px',
                          padding: '12px',
                          color: 'white'
                        }}>
                          <p style={{ fontSize: '11px', opacity: 0.9, margin: '0 0 4px 0', fontWeight: '600' }}>
                            💰 Total Geral
                          </p>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                            <span style={{ opacity: 0.9 }}>Executado:</span>
                            <strong>
                              R$ {totalGeralComNova.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </strong>
                          </div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', marginTop: '2px' }}>
                            <span style={{ opacity: 0.9 }}>Saldo:</span>
                            <strong style={{ color: colors.accent }}>
                              {((totaisComNova.saldoMidia + totaisComNova.saldoProd) >= 0 ? '+' : '')}
                              R$ {(totaisComNova.saldoMidia + totaisComNova.saldoProd).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </strong>
                          </div>
                        </div>
                      </div>

                      <div style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '16px',
                        padding: '12px 16px',
                        background: colors.surface,
                        borderRadius: '10px',
                        border: `1px solid ${colors.border}`
                      }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                            <span style={{ fontSize: '12px', color: colors.textSecondary }}>
                              Taxa de Execução
                            </span>
                            <span style={{ 
                              fontSize: '14px', 
                              fontWeight: '700', 
                              color: status.cor
                            }}>
                              {taxaExecucao.toFixed(1)}%
                            </span>
                          </div>
                          <div style={{
                            width: '100%',
                            height: '6px',
                            background: colors.border,
                            borderRadius: '3px',
                            overflow: 'hidden'
                          }}>
                            <div style={{
                              width: `${Math.min(taxaExecucao, 100)}%`,
                              height: '100%',
                              background: status.cor,
                              borderRadius: '3px',
                              transition: 'width 0.3s ease'
                            }} />
                          </div>
                        </div>
                        <div style={{
                          fontSize: '11px',
                          padding: '4px 12px',
                          borderRadius: '12px',
                          background: status.bg,
                          color: status.cor,
                          fontWeight: '600',
                          whiteSpace: 'nowrap'
                        }}>
                          {status.emoji} {status.status}
                        </div>
                      </div>

                      <div style={{
                        marginTop: '12px',
                        padding: '10px 16px',
                        background: `${colors.primary}08`,
                        borderRadius: '8px',
                        border: `1px dashed ${colors.primary}25`
                      }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: colors.textSecondary }}>
                          <span>
                            📌 {editandoId ? 'Valores atuais:' : 'Valores sem esta campanha:'}
                          </span>
                          <span>
                            Total: R$ {(totaisBase.realizadoMidia + totaisBase.realizadoProd).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                          </span>
                          <span style={{ 
                            color: totalGeralComNova > (totaisBase.realizadoMidia + totaisBase.realizadoProd) 
                              ? colors.danger : colors.success 
                          }}>
                            {editandoId ? '🔄 Nova versão:' : '➕ Com esta campanha:'}
                            R$ {totalGeralComNova.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                          </span>
                        </div>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>

            <div style={{
              display: 'flex',
              justifyContent: 'flex-end',
              gap: '12px',
              marginTop: '32px',
              paddingTop: '24px',
              borderTop: `1px solid ${colors.border}`
            }}>
              <button
                onClick={fecharModal}
                style={{
                  padding: '11px 24px',
                  border: `1.5px solid ${colors.border}`,
                  borderRadius: '10px',
                  background: colors.surface,
                  color: colors.text,
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '14px',
                  transition: 'all 0.2s'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.backgroundColor = colors.background;
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.backgroundColor = colors.surface;
                }}
              >
                Cancelar
              </button>
              <button
                onClick={editandoId ? editarCampanha : adicionarCampanha}
                style={{
                  padding: '11px 24px',
                  background: `linear-gradient(135deg, ${colors.primary} 0%, ${colors.primaryDark} 100%)`,
                  color: colors.surface,
                  border: 'none',
                  borderRadius: '10px',
                  cursor: 'pointer',
                  fontWeight: '600',
                  fontSize: '14px',
                  boxShadow: '0 4px 12px rgba(0, 150, 94, 0.3)',
                  transition: 'all 0.2s'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = 'translateY(-1px)';
                  e.currentTarget.style.boxShadow = '0 6px 16px rgba(0, 150, 94, 0.4)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 4px 12px rgba(0, 150, 94, 0.3)';
                }}
              >
                {editandoId ? 'Salvar Alterações' : 'Criar Campanha'}
              </button>
            </div>
          </div>
        </div>
      )}

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        * {
          box-sizing: border-box;
        }
        
        body {
          margin: 0;
          padding: 0;
        }
        
        @keyframes fadeIn {
          from { opacity: 0; }
          to { opacity: 1; }
        }
        
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
        
        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
        
        input[type="number"] {
          -moz-appearance: textfield;
        }
        
        ::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }
        
        ::-webkit-scrollbar-track {
          background: ${colors.background};
        }
        
        ::-webkit-scrollbar-thumb {
          background: ${colors.border};
          border-radius: 4px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
          background: ${colors.textSecondary};
        }
      `}</style>
    </div>
  );
};

export default App;