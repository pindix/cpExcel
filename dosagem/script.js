// Lógica do Tema Escuro com Memória
const themeBtn = document.getElementById('themeBtn');
const themeIcon = document.getElementById('themeIcon');
const body = document.body;

themeBtn.addEventListener('click', () => {
    if (body.getAttribute('data-theme') === 'dark') {
        body.removeAttribute('data-theme');
        themeIcon.className = 'ri-moon-line';
        localStorage.setItem('tema', 'light'); // Guarda que é light
    } else {
        body.setAttribute('data-theme', 'dark');
        themeIcon.className = 'ri-sun-line';
        localStorage.setItem('tema', 'dark');  // Guarda que é dark
    }
});






// 1. DADOS E CONTROLO
let bancoDados = {}; 
let medAtivo = null;

// 2. ELEMENTOS DE BUSCA E SUGESTÃO
const inputNome = document.getElementById("nome");
const selectPais = document.getElementById("pais");
const divSugestoes = document.getElementById("sugestoes_box");

// 3. CONTENTORES DE CAMPOS (As DIVs para esconder/mostrar)
const camposDivs = {
    peso: document.getElementById("campo_de_peso"),
    idade: document.getElementById("campo_de_idade"),
    dosagem: document.getElementById("campo_de_dosagem"),
    dose: document.getElementById("dose"), // Select direto
    via: document.getElementById("via"),   // Select direto
    intervalo: document.getElementById("intervalo"), // Select diret
    selConcentracao: document.getElementById("selConcentracao")
};

// 4. INPUTS ESPECÍFICOS (Para capturar os valores no cálculo)
const inputs = {
    peso: document.getElementById("peso"),
    idade: document.getElementById("idade"),
    unidadeIdade: document.getElementById("unidade_de_idade"),
    dosagem: document.getElementById("dosagem")
};

// 5. EXIBIÇÃO DE RESULTADOS E UNIDADES
const txtUnidadeDosagem = document.getElementById("unidade_de_dosagem");
const pResultado = document.getElementById("resultado");
















// 1. CARREGAMENTO COM "ANTI-CACHE"
async function carregarDados() {
    try {
        // Adicionamos um número aleatório ao final para forçar o navegador a ler o ficheiro novo
        const response = await fetch('medicamentos.xlsx?v=' + Math.random());
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data);

        bancoDados = {}; // Limpa antes de carregar
        workbook.SheetNames.forEach(nome => {
            bancoDados[nome.toLowerCase().trim()] = XLSX.utils.sheet_to_json(workbook.Sheets[nome]);
        });
        console.log("Base de dados atualizada e carregada.");
    } catch (e) {
        console.error("Erro ao carregar o ficheiro Excel.", e);
    }
}

// 2. FUNÇÃO DE SUGESTÕES (COM LIMPEZA DE ESPAÇOS)
function gerirSugestoes() {
    const termo = inputNome.value.trim();
    const termoLower = termo.toLowerCase();
    const paisAtivo = selectPais.value.toLowerCase();

    if (!termo) {
        divSugestoes.style.display = "none";
        return;
    }

    // Busca nas abas (País e Universal)
    let basePais = bancoDados[paisAtivo] || [];
    let baseUniversal = bancoDados["universal"] || [];
    let baseTotal = [...basePais, ...baseUniversal];

    // FILTRO REVISADO: Limpa espaços de cada nome do Excel antes de comparar
    const nomesFiltrados = [...new Set(
        baseTotal
            .map(m => String(m.nome).trim()) // Remove espaços invisíveis do Excel
            .filter(nome => nome.toLowerCase().includes(termoLower))
    )];

    // Lógica de fechamento
    if (nomesFiltrados.length === 0) {
    divSugestoes.style.display = "none";
    return;
}

    divSugestoes.innerHTML = "";
    divSugestoes.style.display = "block";

    nomesFiltrados.slice(0, 6).forEach(nome => {
        const item = document.createElement("div");
        item.className = "sugestao_item";
        item.innerText = nome;
        
        item.onclick = () => {
            inputNome.value = nome;
            divSugestoes.style.display = "none";
            escolherLinha('silencioso');
            exibirCampos();
        };
        divSugestoes.appendChild(item);
    });
}



// --- 1. FUNÇÃO ESCOLHER LINHA (O CÉREBRO) ---

// --- 1. FUNÇÃO ESCOLHER LINHA (O CÉREBRO) ---
function escolherLinha(modo) {
    
    const nome = inputNome.value.trim().toLowerCase();
    const pais = selectPais.value.toLowerCase();
    
    if (!nome) { medAtivo = null; return; }

    // Busca todas as linhas do medicamento (País > Universal)
    let todasAsLinhas = bancoDados[pais]?.filter(m => String(m.nome).toLowerCase() === nome) || [];
    if (todasAsLinhas.length === 0) {
        todasAsLinhas = bancoDados["universal"]?.filter(m => String(m.nome).toLowerCase() === nome) || [];
    }

    if (todasAsLinhas.length === 0) { medAtivo = null; return; }

    // Se for o primeiro carregamento (silencioso), pegamos a primeira linha e saímos
    if (modo === 'silencioso') {
        medAtivo = todasAsLinhas[0];
        pResultado.innerText = ""; // Sem feedback ao digitar
        return;
    }

    // Se for 'ajuste' (usuário mexeu nos campos), filtramos por proximidade
    let filtradas = todasAsLinhas;
    const doseSel = camposDivs.dose.value;
    const viaSel = camposDivs.via.value;

    // Tenta filtrar por Dose e Via
    let tentativaDoseVia = filtradas.filter(m => m.dose === doseSel && m.via === viaSel);
    if (tentativaDoseVia.length > 0) filtradas = tentativaDoseVia;

    // Tenta filtrar por Peso/Idade
    const pVal = parseFloat(inputs.peso.value);
    const iVal = parseFloat(inputs.idade.value);
    let tentativaBio = filtradas.filter(m => {
        const batePeso = !pVal || (pVal >= m.peso_minimo && pVal <= m.peso_maximo);
        const bateIdade = !iVal || (iVal >= m.idade_minima && iVal <= m.idade_maxima);
        return batePeso && bateIdade;
    });
    if (tentativaBio.length > 0) filtradas = tentativaBio;

    medAtivo = filtradas[0];

    // Geração do Feedback (Apenas no modo ajuste)
    if (medAtivo.dose === doseSel && medAtivo.via === viaSel) {
        pResultado.innerText = "✅ Configuração exata.";
        pResultado.style.background = "var(--primary)";
        pResultado.style.textAlign = "center";
    } else {
        pResultado.innerText = `⚠️ Tivemos que ajustarpara: ${medAtivo.dose} | ${medAtivo.via.toUpperCase()}`;
        pResultado.style.background = "orange";
        pResultado.style.textAlign = "left";
    }
}

// --- 2. FUNÇÃO EXIBIR CAMPOS (O VISUAL) ---
function exibirCampos() {
    if (!medAtivo) {
        Object.values(camposDivs).forEach(div => div.style.display = "none");
        inputs.peso.value = "";
        inputs.idade.value = "";
        inputs.dosagem.value = "";
        pResultado.innerHTML = "";
        pResultado.style.background = "var(--primary)";
        return;
    }

    // 1. Visibilidade dos campos básicos (Peso, Idade, Dosagem, etc)
    const c = medAtivo.campos ? medAtivo.campos.toLowerCase() : "";
    camposDivs.peso.style.display    = c.includes("peso") ? "flex" : "none";
    camposDivs.idade.style.display   = c.includes("idade") ? "flex" : "none";
    camposDivs.dosagem.style.display = c.includes("dosagem") ? "flex" : "none";
    camposDivs.dose.style.display    = c.includes("dose") ? "flex" : "none";
    camposDivs.via.style.display     = c.includes("via") ? "flex" : "none";
    
    // 2. Lógica de Concentração Dinâmica (O nosso Select)
    // Verificamos se a célula de concentração contém o separador "|"
    const concRaw = String(medAtivo.concentracao || "");
    
    if (concRaw.includes("|")) {
        camposDivs.selConcentracao.style.display = "block"; // Mostra o select
        camposDivs.selConcentracao.innerHTML = ""; // Limpa opções anteriores

        // Divide: "4 mg/ml|4; 4 mg/2ml|2" -> ["4 mg/ml|4", " 4 mg/2ml|2"]
        const grupos = concRaw.split(";");
        
        grupos.forEach(grupo => {
            const partes = grupo.split("|");
            if (partes.length === 2) {
                const opt = document.createElement("option");
                opt.innerText = partes[0].trim(); // O que o usuário vê
                opt.value = partes[1].trim();     // O número que o cálculo usará
                camposDivs.selConcentracao.appendChild(opt);
            }
        });
    } else {
        // Se for um valor simples (ex: "500"), esconde o select
        camposDivs.selConcentracao.style.display = "none";
    }

    // 3. Unidades e Intervalos
    txtUnidadeDosagem.innerText = medAtivo.unidade_dosagem || "mg/kg";

    camposDivs.intervalo.innerHTML = "";
    if (medAtivo.intervalo) {
        camposDivs.intervalo.style.display = "flex";
        const opcoes = String(medAtivo.intervalo).split(",");
        opcoes.forEach(opt => {
            const horas = parseInt(opt.trim());
            if (!isNaN(horas)) {
                const elemento = document.createElement("option");
                elemento.value = 24 / horas;
                elemento.innerText = `${horas}/${horas}h`;
                camposDivs.intervalo.appendChild(elemento);
            }
        });
    } else {
        camposDivs.intervalo.style.display = "none";
    }
}

// --- 3. ATUALIZAÇÃO DOS GATILHOS (IMPORTANTE) ---

// Nas sugestões ou ao terminar de escrever:
// escolherLinha('silencioso'); exibirCampos();

// Nos inputs de Peso, Dose, Via:
// escolherLinha('ajuste'); exibirCampos();












function calcular() {
    // 1. Animação de Feedback (Corrigido para usar pResultado)
    pResultado.classList.remove("vibrar");
    void pResultado.offsetWidth; // Força o re-flow
    pResultado.classList.add("vibrar");
    
    // ... restante do código
    
    
    // 1. Verificação de existência do medicamento
    if (!medAtivo) {
        pResultado.innerHTML = "Medicamento não encontrado!";
        pResultado.style.background = "red";
        pResultado.style.textAlign = "center";
        return;
    }

    // 2. Captura de valores dos inputs e do medAtivo
    let peso = parseFloat(inputs.peso.value) || 0;
    let idade = parseFloat(inputs.idade.value) || 0;
    let dosagem = parseFloat(inputs.dosagem.value) || 0;
    let intervalo = parseFloat(camposDivs.intervalo.value) || 1;
    
    let concentracao = 1; 

// Se o select de concentração estiver visível, pegamos o VALOR selecionado nele
if (camposDivs.selConcentracao && camposDivs.selConcentracao.style.display !== "none") {
    concentracao = parseFloat(camposDivs.selConcentracao.value) || 1;
} else {
    // Caso contrário (medicamento simples), pegamos o valor direto da coluna do Excel
    concentracao = parseFloat(medAtivo.concentracao) || 1;
}

    // 3. GUARDIAN: Correção e Alertas Simplificados
    

// 1. SEGURANÇA: DOSAGEM
    if (camposDivs.dosagem.style.display !== "none" && inputs.dosagem.value !== "") {
        const n = parseFloat(medAtivo.dosagem_minima) || 0;
        const m = parseFloat(medAtivo.dosagem_maxima) || Infinity;
        if (dosagem < n || dosagem > m) {
            dosagem = dosagem < n ? n : m;
            alert(`⚠️Sistema de Segurança.\n\nDosagem corrigida para ${dosagem}.\nDeve estar entre ${n}-${m}.`);
            inputs.dosagem.value = dosagem;
        }
    }

    // 2. SEGURANÇA: PESO
    if (camposDivs.peso.style.display !== "none" && inputs.peso.value !== "") {
        const n = parseFloat(medAtivo.peso_minimo) || 0;
        const m = parseFloat(medAtivo.peso_maximo) || Infinity;
        if (peso < n || peso > m) {
            peso = peso < n ? n : m;
            alert(`⚠️Sistema de Segurança.\n\nPeso corrigido para ${peso}kg.\nDeve estar entre ${n}-${m}kg.`);
            inputs.peso.value = peso;
        }
    }

    // 3. SEGURANÇA: IDADE
    if (camposDivs.idade.style.display !== "none" && inputs.idade.value !== "") {
        const n = parseFloat(medAtivo.idade_minima) || 0;
        const m = parseFloat(medAtivo.idade_maxima) || Infinity;
        if (idade < n || idade > m) {
            idade = idade < n ? n : m;
            alert(`⚠️Sistema de Segurança.\n\nIdade corrigida para ${idade}.\nDeve estar entre ${n}-${m}.`);
            inputs.idade.value = idade;
        }
    }
    
    
    
    // 4. PROCESSAMENTO DA FÓRMULA
    try {
let formulaTexto = medAtivo.formula;

// 1. Substituições dos valores usando o prefixo #
// Usamos \b para garantir que #p não mude palavras que apenas contenham a letra p
let formulaProcessada = formulaTexto
    .replace(/#p\b/g, peso)
    .replace(/#d\b/g, dosagem)
    .replace(/#c\b/g, concentracao)
    .replace(/#i\b/g, intervalo);

// 2. RESOLVER A MATEMÁTICA (O que está dentro de { })
// O eval processa os números que já foram substituídos acima
let resultadoCalculado = formulaProcessada.replace(/{([^}]+)}/g, (match, expressao) => {
    try {
        // Resolve a conta e limita a 2 casas decimais
        let valor = eval(expressao);
        return isNaN(valor) ? "Erro" : valor.toFixed(1);
    } catch (e) {
        return "Erro"; 
    }
});
// 3. APLICAR O ESTILO DE DOSE (O que está dentro de [ ])
// Remove os [ ] e envolve o conteúdo na classe CSS de destaque
let resultadoFinal = resultadoCalculado.replace(/\[([^\]]+)\]/g, (match, conteudo) => {
    return `<span class="dose-destaque">${conteudo}</span>`;
});
        
        
        
        
        
        
        

// 4. EXIBIR NO HTML
// Aqui o símbolo # sozinho (fora dos tokens) funciona como quebra de linha <br>
pResultado.innerHTML = resultadoFinal.replace(/#/g, "<br>");
        pResultado.style.color = "white";
        pResultado.style.background = "var(--primary)";
        pResultado.style.textAlign = "left";

    } catch (erro) {
        console.error("Erro no cálculo:", erro);
        pResultado.innerHTML = "⚠️ Erro na fórmula da base de dados!.";
    }
}



function mostrarAjudaAjuste() {
    alert("O sistema selecionou a melhor opção baseada nos critérios de segurança (Peso, Idade ou Via) definidos para este medicamento.");
}

function limpar() {
    // 1. Limpa os textos escritos
    inputNome.value = "";
    inputs.peso.value = "";
    inputs.idade.value = "";
    inputs.dosagem.value = "";
    
    // 2. Reseta a lógica
    medAtivo = null;
    
    // 3. Esconde os campos e limpa o resultado
    exibirCampos();
    pResultado.innerHTML = "";
    pResultado.style.background = "var(--primary)";
    
    pResultado.classList.remove("vibrar");
    void pResultado.offsetWidth; // Força o re-flow
    pResultado.classList.add("vibrar");
}


selectPais.addEventListener('change', () => {
    localStorage.setItem('pais', selectPais.value);
    const nomePais = selectPais.options[selectPais.selectedIndex].text;
    
    // 1. Limpa tudo o que estava no ecrã antes
    limpar();
    
    // 2. Feedback de "A carregar"
    pResultado.innerHTML = "A carregar...";
    pResultado.style.background = "orange";
    pResultado.style.color = "white";
    pResultado.style.textAlign = "left";
    
    // 3. Simulação de atualização (Timeout)
    setTimeout(() => {
        pResultado.innerHTML = `Padrões de <b>${nomePais}</b> carregados. <br> Tudo pronto.`;
        pResultado.style.background = "var(--primary)"; // Usa a cor principal do teu CSS
    }, 1200);
});






// B. No clique da sugestão (Gatilho Silencioso)
// Nota: Ajusta a tua função gerirSugestoes para que o item.onclick chame isto:
/*
item.onclick = () => {
    inputNome.value = nome;
    divSugestoes.style.display = "none";
    escolherLinha('silencioso');
    exibirCampos();
};
*/


inputNome.addEventListener("input", () => {
    const termo = inputNome.value.trim();

    if (termo.length > 0) {
        // 1. Tenta encontrar o medicamento exato primeiro
        escolherLinha('silencioso');

        if (medAtivo) {
            // Se achou o nome exato (ex: terminou de escrever)
            divSugestoes.style.display = "none"; // Fecha sugestões
            exibirCampos(); // Mostra os campos
        } else {
            // Se ainda não é um nome exato, abre as sugestões
            gerirSugestoes(); 
            exibirCampos(); // Vai esconder os campos (correto, pois o nome está incompleto)
        }
    } else {
        // Se apagou tudo
        medAtivo = null;
        divSugestoes.style.display = "none";
        exibirCampos();
    }
});

// C. Nos campos de Dados (Gatilho de Ajuste/Feedback)
// Sempre que o utilizador mexer nestes campos, o sistema verifica se a linha ainda é a melhor
const gatilhosAjuste = [
    camposDivs.dose, 
    camposDivs.via, 
    inputs.peso, 
    inputs.idade, 
    inputs.dosagem,
];

gatilhosAjuste.forEach(elemento => {
    // Usamos 'change' para selects e 'input' para campos de número
    const evento = elemento.tagName === 'SELECT' ? 'change' : 'input';
    
    elemento.addEventListener(evento, () => {
        // Só faz sentido ajustar se já houver um medicamento selecionado
        if (medAtivo) {
            escolherLinha('ajuste'); 
            exibirCampos();
        }
    });
});


// Gatilho para cálculos instantâneos (apenas se os dados necessários existirem)
[camposDivs.selConcentracao, camposDivs.intervalo].forEach(seletor => {
    seletor.addEventListener('change', () => {
        if (!medAtivo) return;

        // Verificamos se os campos que NÃO estão escondidos foram preenchidos
        const pesoOK = camposDivs.peso.style.display === "none" || inputs.peso.value.trim() !== "";
        const idadeOK = camposDivs.idade.style.display === "none" || inputs.idade.value.trim() !== "";
        const dosagemOK = camposDivs.dosagem.style.display === "none" || inputs.dosagem.value.trim() !== "";

        // Só calcula se todos os campos necessários para este medicamento tiverem dados
        if (pesoOK && idadeOK && dosagemOK) {
            calcular();
        }
    });
});





document.addEventListener("click", (e) => {
    if (e.target !== inputNome && e.target !== divSugestoes) {
        divSugestoes.style.display = "none";
    }
});




// Quando a página carrega, recupera as escolhas anonimamente
window.addEventListener('load', () => {
    // Restaurar Tema
    const temaSalvo = localStorage.getItem('tema');
    if (temaSalvo === 'dark') {
        body.setAttribute('data-theme', 'dark');
        themeIcon.className = 'ri-sun-line';
    }

    // Restaurar País
    const paisSalvo = localStorage.getItem('pais');
    if (paisSalvo) {
        selectPais.value = paisSalvo;
    }
    
    // Inicia o carregamento do Excel
    carregarDados();
});


