# -*- coding: utf-8 -*-
"""
Descrição_EFPC.py
==============
Base de "fichas" (descrição + patrocinadoras) das EFPCs, para alimentar o
modal "Ficha da EFPC" do dashboard.

FONTE: pesquisa web (sites institucionais das próprias fundações, Previc,
imprensa especializada — Investidor Institucional, Teletime — e registros
públicos de CNPJ), realizada em jul/2026. Patrocinadores de fundos de pensão
podem mudar (fusões, retirada de patrocínio, privatizações), então vale
reconferir periodicamente — principalmente casos em situação especial citados
abaixo (ex: Oi em recuperação judicial, Fachesf avaliando incorporação).

Cobertura atual: as 24 EFPCs classificadas como Tier F3 = CONSULTING.
Estrutura pronta para receber as ~73 de OCIO I / OCIO II na sequência —
basta adicionar novas entradas neste mesmo dicionário, chave = NU_MATRICULA_EFPC.

Campos de cada ficha:
  patrocinadoras : list[str]  -> nome(s) da(s) empresa(s)/entidade(s) patrocinadora(s)
  setor          : str        -> setor econômico predominante
  descricao      : str        -> 2-3 frases de contexto (história/porte/situação)
  fonte          : str        -> observação curta sobre a fonte/confiabilidade do dado
"""

FICHAS_EFPC = {
 
    # ══════════════════════════════════════════
    # TIER F3 = CONSULTING (24)
    # ══════════════════════════════════════════
 
    93: {  # BANESPREV
        "patrocinadoras": ["Banco Santander (Brasil) S.A. (sucessor do Banespa)"],
        "setor": "Financeiro / Bancário",
        "descricao": "Fundo de pensão criado para os funcionários do extinto Banespa "
                     "(Banco do Estado de São Paulo). Após a privatização e incorporação "
                     "do Banespa pelo Santander em 2000, o Santander assumiu a condição de "
                     "patrocinador junto à entidade.",
        "fonte": "Confirmado - registros públicos e histórico do setor bancário paulista.",
    },
 
    3188: {  # BB PREVIDENCIA
        "patrocinadoras": ["Banco do Brasil S.A."],
        "setor": "Financeiro / Bancário",
        "descricao": "Fundo de pensão fechado do Banco do Brasil, distinto da Previ "
                     "(que atende o quadro histórico do banco). Administra planos "
                     "complementares de aposentadoria para empregados do BB.",
        "fonte": "Confirmado - nome/razão social já indica o patrocinador.",
    },
 
    237: {  # CERES
        "patrocinadoras": ["Embrapa", "Emater-MG", "Epamig", "Epagri", "Cidasc", "ABDI", "Emater-DF"],
        "setor": "Público / Pesquisa agropecuária",
        "descricao": "Entidade multipatrocinada criada em 1979 pela Embrapa e pela extinta "
                     "Embrater para atrair e reter profissionais do setor de pesquisa e "
                     "extensão rural. Hoje reúne 7 patrocinadoras ligadas ao agronegócio "
                     "público, com cerca de 11,5 mil participantes e 8 mil assistidos.",
        "fonte": "Confirmado via site institucional e LinkedIn oficial da Ceres Previdência.",
    },
 
    312: {  # ECONOMUS
        "patrocinadoras": ["Banco do Brasil S.A. (sucessor do Banco Nossa Caixa)"],
        "setor": "Financeiro / Bancário",
        "descricao": "Instituto de seguridade criado em 1977 como política de RH do "
                     "extinto Banco Nossa Caixa (BNC). Com a aquisição do BNC pelo Banco "
                     "do Brasil em 2009, o BB passou a ser o patrocinador da entidade, "
                     "que também administra assistência médico-hospitalar.",
        "fonte": "Confirmado - site institucional Economus.",
    },
 
    361: {  # FACHESF
        "patrocinadoras": ["Chesf - Companhia Hidro Elétrica do São Francisco (grupo Eletrobras)"],
        "setor": "Energia elétrica",
        "descricao": "Criada em 1972 pela Chesf, é um dos maiores fundos de pensão do "
                     "Norte/Nordeste, com patrimônio superior a R$ 5,5 bi. Está entre as "
                     "EFPCs do setor elétrico (junto com Eletros, Real Grandeza, Previnorte "
                     "e Elos) que discutem uma possível unificação em nova entidade após a "
                     "privatização da Eletrobras.",
        "fonte": "Confirmado - site institucional e reportagem Investidor Institucional (2026).",
    },
 
    391: {  # FAPES
        "patrocinadoras": ["BNDES - Banco Nacional de Desenvolvimento Econômico e Social"],
        "setor": "Público / Financeiro (banco de fomento)",
        "descricao": "Fundo de pensão dos empregados do BNDES, uma das maiores EFPCs "
                     "ligadas a instituições públicas federais de fomento.",
        "fonte": "Alta confiança - nome/razão social indica o patrocinador; recomenda-se confirmar dados atuais no site oficial.",
    },
 
    4203: {  # FATL
        "patrocinadoras": ["Oi S.A. (ex-Telemar, sistema Telebras)"],
        "setor": "Telecomunicações",
        "descricao": "Fundação multipatrocinada (Rio de Janeiro) cuja principal "
                     "patrocinadora é a Oi. Administra os planos PBS-Telemar, TelemarPrev, "
                     "TCSPREV e outros. Está sob atenção por causa da recuperação judicial "
                     "da Oi, com associações de participantes pedindo à Previc que garanta "
                     "a autonomia da fundação frente à gestão judicial da patrocinadora.",
        "fonte": "Confirmado - site institucional e reportagens TeleTime/Telesíntese (2026).",
    },
 
    1479: {  # FORLUZ
        "patrocinadoras": ["Cemig - Companhia Energética de Minas Gerais"],
        "setor": "Energia elétrica",
        "descricao": "Fundo de pensão dos empregados da Cemig, um dos maiores fundos "
                     "ligados ao setor elétrico mineiro.",
        "fonte": "Alta confiança - nome/razão social ('Forluminas') indica o patrocinador; recomenda-se confirmar no site oficial.",
    },
 
    1523: {  # FUNCEF
        "patrocinadoras": ["Caixa Econômica Federal"],
        "setor": "Financeiro / Bancário (público)",
        "descricao": "Um dos maiores fundos de pensão do Brasil, fundo dos empregados "
                     "da Caixa Econômica Federal.",
        "fonte": "Confirmado - amplamente documentado (um dos maiores fundos do país).",
    },
 
    1239: {  # FUNCESP (Vivest)
        "patrocinadoras": ["Grupo AES (AES Eletropaulo / AES Tietê)", "Grupo CPFL", "CTEEP",
                            "Duke Energy", "Elektro", "EMAE", "CESP (Auren Energia)"],
        "setor": "Energia elétrica (multipatrocinado)",
        "descricao": "Também conhecida pela marca Vivest, é o maior fundo de pensão "
                     "patrocinado por empresas privadas do país e a 4ª maior EFPC em "
                     "ativos (~R$ 23 bi), com mais de 100 mil participantes. Nasceu como "
                     "fundo da CESP e hoje reúne diversas empresas do setor elétrico "
                     "paulista que sucederam a estatal privatizada.",
        "fonte": "Confirmado - reportagem Investidor Institucional e site institucional Funcesp/Vivest.",
    },
 
    285: {  # FUNDACAO COPEL
        "patrocinadoras": ["Copel - Companhia Paranaense de Energia (grupo Copel, 7 patrocinadoras)"],
        "setor": "Energia elétrica",
        "descricao": "Maior previdência privada do Sul do país, fundada em 1971. Reúne "
                     "7 empresas patrocinadoras e 3 instituidoras do grupo Copel, com mais "
                     "de R$ 15 bi em recursos administrados e 20 mil participantes.",
        "fonte": "Confirmado - site institucional Fundação Copel.",
    },
 
    4724: {  # FUNPRESP-EXE
        "patrocinadoras": ["União (servidores públicos federais do Poder Executivo)"],
        "setor": "Público / Administração federal",
        "descricao": "Fundo de previdência complementar dos servidores públicos federais "
                     "do Poder Executivo, criado pela Lei 12.618/2012. Não tem patrocinador "
                     "privado - é vinculado ao regime de previdência complementar do "
                     "funcionalismo público federal.",
        "fonte": "Confirmado - natureza pública amplamente documentada (Lei 12.618/2012).",
    },
 
    611: {  # ITAU UNIBANCO
        "patrocinadoras": ["Itaú Unibanco S.A."],
        "setor": "Financeiro / Bancário",
        "descricao": "Fundo de previdência complementar fechado dos empregados do "
                     "conglomerado Itaú Unibanco, um dos maiores bancos privados do país.",
        "fonte": "Confirmado - razão social indica diretamente o patrocinador.",
    },
 
    1482: {  # MULTIBRA
        "patrocinadoras": ["Grupo Bradesco (ex-fundo do HSBC, incorporado em 2016)"],
        "setor": "Financeiro (multipatrocinado)",
        "descricao": "Fundo multipatrocinado originado do antigo fundo do HSBC no Brasil; "
                     "após a venda do HSBC ao Bradesco, passou a ser administrado pela "
                     "Kirton (grupo Bradesco Seguros) e renomeado para Multibra. Junto com "
                     "o Multipensions (também do Bradesco), forma um dos maiores "
                     "multipatrocinados do mercado.",
        "fonte": "Confirmado - reportagem Investidor Institucional e processos públicos (Jusbrasil).",
    },
 
    2258: {  # MULTIPREV
        "patrocinadoras": ["Multipatrocinado - patrocinadoras específicas a confirmar"],
        "setor": "Multipatrocinado",
        "descricao": "Um dos grandes fundos multipatrocinados do mercado brasileiro, "
                     "citado como concorrente direto do Multibra e do Icatu no segmento. "
                     "Lista detalhada de empresas patrocinadoras ainda não confirmada "
                     "com fonte primária.",
        "fonte": "Baixa confiança - recomenda-se pesquisar diretamente no site da entidade ou cadastro Previc antes de usar em prospecção.",
    },
 
    655: {  # PETROS
        "patrocinadoras": ["Petrobras"],
        "setor": "Petróleo e gás (estatal)",
        "descricao": "Um dos maiores fundos de pensão da América Latina, fundo dos "
                     "empregados da Petrobras e de empresas do sistema Petrobras.",
        "fonte": "Confirmado - amplamente documentado.",
    },
 
    691: {  # POSTALIS
        "patrocinadoras": ["Correios - Empresa Brasileira de Correios e Telégrafos (ECT)"],
        "setor": "Público / Serviços postais",
        "descricao": "Fundo de pensão dos empregados dos Correios, que passou por um "
                     "processo de reestruturação após déficit atuarial identificado na "
                     "década de 2010.",
        "fonte": "Confirmado - amplamente documentado na imprensa especializada.",
    },
 
    1781: {  # PREVI/BB
        "patrocinadoras": ["Banco do Brasil S.A."],
        "setor": "Financeiro / Bancário",
        "descricao": "A Previ é o maior fundo de pensão da América Latina, fundo "
                     "histórico dos funcionários do Banco do Brasil.",
        "fonte": "Confirmado - amplamente documentado (maior EFPC do país).",
    },
 
    1033: {  # PREVIDÊNCIA USIMINAS
        "patrocinadoras": ["Usiminas - Usinas Siderúrgicas de Minas Gerais"],
        "setor": "Siderurgia",
        "descricao": "Fundo de pensão dos empregados da Usiminas.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    864: {  # REAL GRANDEZA
        "patrocinadoras": ["Furnas Centrais Elétricas", "Eletronuclear (grupo Eletrobras)"],
        "setor": "Energia elétrica",
        "descricao": "Fundo de pensão do setor elétrico ligado a Furnas e à Eletronuclear, "
                     "hoje parte do grupo de EFPCs do sistema Eletrobras (junto com Eletros, "
                     "Fachesf, Previnorte e Elos) que avalia um projeto de unificação em "
                     "nova entidade.",
        "fonte": "Confirmado - reportagem Investidor Institucional (2026) e FNU.",
    },
 
    881: {  # REFER
        "patrocinadoras": ["Rede Ferroviária Federal S.A. - RFFSA (extinta em 1999, sucedida pela União)"],
        "setor": "Público / Ferroviário (herança)",
        "descricao": "Fundo criado para os ferroviários da extinta RFFSA, liquidada em "
                     "1999; hoje suas obrigações remanescentes têm acompanhamento da União "
                     "como sucessora legal.",
        "fonte": "Média confiança - contexto histórico conhecido; recomenda-se confirmar situação atual de patrocínio no cadastro Previc.",
    },
 
    967: {  # SISTEL
        "patrocinadoras": ["Sistema Telebras (operadoras de telecom sucessoras, principalmente Oi)"],
        "setor": "Telecomunicações",
        "descricao": "Uma das mais antigas e tradicionais fundações de previdência do "
                     "setor de telecomunicações, originada do extinto Sistema Telebras. "
                     "Está no mesmo grupo de fundações do setor (junto com Telos, FATL e "
                     "Visão Prev) que compartilham histórico de decisões judiciais sobre "
                     "expurgos inflacionários e imposto de renda.",
        "fonte": "Confirmado - contexto histórico e menções cruzadas com FATL/Telos em fontes jurídicas.",
    },
 
    998: {  # TELOS
        "patrocinadoras": ["Embratel (grupo Claro / América Móvil)"],
        "setor": "Telecomunicações",
        "descricao": "Fundação de seguridade social ligada historicamente à Embratel, "
                     "parte do grupo de fundações do setor de telecomunicações que "
                     "sucederam o Sistema Telebras.",
        "fonte": "Média confiança - nome ('Fundação Embratel') indica o patrocinador histórico; confirmar situação atual (pós fusão Embratel/Claro).",
    },
 
    2083: {  # VALIA
        "patrocinadoras": ["Vale S.A."],
        "setor": "Mineração",
        "descricao": "Fundo de pensão dos empregados da Vale (ex-Companhia Vale do Rio "
                     "Doce), uma das maiores EFPCs do país.",
        "fonte": "Confirmado - amplamente documentado.",
    },
 
    # ══════════════════════════════════════════
    # TIER F3 = OCIO I (18)
    # ══════════════════════════════════════════
 
    117: {  # BANRISUL/FBSS
        "patrocinadoras": ["Banrisul - Banco do Estado do Rio Grande do Sul"],
        "setor": "Financeiro / Bancário (público estadual)",
        "descricao": "Fundo de pensão dos empregados do Banrisul, banco estadual do "
                     "Rio Grande do Sul.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    182: {  # CAPEF
        "patrocinadoras": ["Banco do Nordeste do Brasil S.A. (BNB)"],
        "setor": "Financeiro / Bancário (público federal)",
        "descricao": "Fundo de pensão dos funcionários do Banco do Nordeste do Brasil.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    1208: {  # CBS
        "patrocinadoras": ["CSN - Companhia Siderúrgica Nacional", "CSN Mineração/Namisa", "CSN Cimentos"],
        "setor": "Siderurgia / Mineração",
        "descricao": "Fundada em 1960, é o 5º fundo de pensão mais antigo do Brasil. "
                     "Multipatrocinada pelo grupo CSN, incluindo braços de mineração e "
                     "cimento do conglomerado.",
        "fonte": "Confirmado - site institucional CBS Previdência.",
    },
 
    223: {  # CENTRUS
        "patrocinadoras": ["Banco Central do Brasil"],
        "setor": "Público / Autoridade monetária",
        "descricao": "Fundo de previdência complementar dos servidores do Banco Central "
                     "do Brasil.",
        "fonte": "Alta confiança - razão social ('Banco Central de Previdência Privada') indica diretamente o patrocinador.",
    },
 
    326: {  # ELETROS
        "patrocinadoras": ["Eletrobras (holding)"],
        "setor": "Energia elétrica",
        "descricao": "Fundo de pensão dos empregados da própria Eletrobras (holding), "
                     "parte do grupo de EFPCs do setor elétrico (com Fachesf, Real "
                     "Grandeza, Previnorte e Elos) que discute uma possível unificação em "
                     "nova entidade após a privatização da estatal em 2022.",
        "fonte": "Confirmado - reportagem Investidor Institucional/FNU (2024/2026).",
    },
 
    4581: {  # EMBRAER PREV
        "patrocinadoras": ["Embraer - Empresa Brasileira de Aeronáutica"],
        "setor": "Aeroespacial / Defesa",
        "descricao": "Fundo de pensão dos empregados da Embraer.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    1081: {  # FAMILIA PREVIDENCIA (Fundação CEEE / ELETROCEEE)
        "patrocinadoras": ["Grupo CEEE - Companhia Estadual de Energia Elétrica", "AES Sul Distribuidora Gaúcha",
                            "Rio Grande Energia S/A", "CGTEE - Cia de Geração Térmica de Energia Elétrica",
                            "CRM - Cia Riograndense de Mineração"],
        "setor": "Energia elétrica (multipatrocinado, RS)",
        "descricao": "Criada em 1979 como Fundação CEEE, hoje opera com a marca Família "
                     "Previdência (razão social ainda é 'Fundação CEEE de Seguridade "
                     "Social - ELETROCEEE'). É o maior fundo de pensão do Rio Grande do "
                     "Sul, com patrimônio superior a R$ 7 bi, multipatrocinado desde as "
                     "privatizações da distribuição de energia no estado.",
        "fonte": "Confirmado - LinkedIn oficial e site institucional Família Previdência.",
    },
 
    571: {  # FIBRA
        "patrocinadoras": ["Itaipu Binacional"],
        "setor": "Energia elétrica (binacional Brasil-Paraguai)",
        "descricao": "Fundo de previdência e assistência dos empregados brasileiros da "
                     "Itaipu Binacional.",
        "fonte": "Alta confiança - razão social ('Fundação Itaipu BR') indica diretamente o patrocinador.",
    },
 
    504: {  # FUNBEP
        "patrocinadoras": ["Itaú Unibanco S.A.", "Banco Itaú BBA", "Banco Itaú Leasing",
                            "Fundação Itaú Unibanco", "TECPAR - Instituto de Tecnologia do Paraná"],
        "setor": "Financeiro (multipatrocinado)",
        "descricao": "Originado do antigo Banestado (Banco do Estado do Paraná), "
                     "incorporado pelo Itaú em 2000. Hoje é multipatrocinado pelo grupo "
                     "Itaú Unibanco, com sede em Curitiba, administrando cerca de 6 mil "
                     "participantes/assistidos.",
        "fonte": "Confirmado - relatório anual e site institucional Funbep.",
    },
 
    4741: {  # FUNPRESP-JUD
        "patrocinadoras": ["União (servidores públicos federais do Poder Judiciário)"],
        "setor": "Público / Poder Judiciário federal",
        "descricao": "Fundo de previdência complementar dos servidores do Poder "
                     "Judiciário da União, análogo à Funpresp-Exe (que atende o Executivo).",
        "fonte": "Alta confiança - natureza pública análoga à Funpresp-Exe, amplamente documentada.",
    },
 
    1571: {  # IBM
        "patrocinadoras": ["IBM Brasil"],
        "setor": "Tecnologia",
        "descricao": "Fundo de previdência complementar dos empregados da IBM no Brasil.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    3126: {  # IFM
        "patrocinadoras": ["Multipatrocinado - entre outros: Pfizer", "Novartis/Sandoz",
                            "planos do grupo EDP Energias do Brasil (ex-Enerprev)"],
        "setor": "Multipatrocinado (plataforma aberta)",
        "descricao": "O Itajubá Fundo Multipatrocinado é uma plataforma que administra "
                     "planos previdenciários de diversas empresas de portes e setores "
                     "distintos (farmacêutico, energia etc.), funcionando como "
                     "'multi-empregador' para empresas que não têm fundo próprio.",
        "fonte": "Confirmado - site institucional e reportagem Investidor Institucional (transferência de planos Enerprev).",
    },
 
    1852: {  # PREVI-GM
        "patrocinadoras": ["General Motors do Brasil"],
        "setor": "Automotivo",
        "descricao": "Fundo de pensão dos empregados da General Motors do Brasil.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    4279: {  # QUANTA
        "patrocinadoras": ["Sistema cooperativista - patrocinadoras específicas a confirmar"],
        "setor": "Cooperativismo (multipatrocinado)",
        "descricao": "Entidade de previdência complementar de Florianópolis (SC), com "
                     "DNA cooperativista e forte atuação em educação financeira; "
                     "administra o plano Prevcoop para mais de 190 mil participantes. "
                     "Lista detalhada de cooperativas/empresas patrocinadoras não "
                     "confirmada com fonte primária.",
        "fonte": "Baixa confiança para a lista de patrocinadoras - recomenda-se checar direto com a entidade.",
    },
 
    2511: {  # SANTANDERPREVI
        "patrocinadoras": ["Banco Santander (Brasil) S.A."],
        "setor": "Financeiro / Bancário",
        "descricao": "Fundo de previdência complementar dos empregados do Santander no "
                     "Brasil.",
        "fonte": "Alta confiança - razão social indica diretamente o patrocinador.",
    },
 
    941: {  # SERPROS
        "patrocinadoras": ["Serpro - Serviço Federal de Processamento de Dados"],
        "setor": "Público / Tecnologia da informação (federal)",
        "descricao": "Fundo multipatrocinado ligado ao Serpro, empresa pública federal "
                     "de TI, com patrimônio superior a R$ 7,7 bi.",
        "fonte": "Alta confiança - nome e histórico associam a entidade ao Serpro; confirmar outras patrocinadoras no site oficial.",
    },
 
    3174: {  # VEXTY
        "patrocinadoras": ["Novonor (ex-Odebrecht) - patrocinadora-fundadora", "mais de 110 empresas-patrocinadoras adicionais"],
        "setor": "Multipatrocinado (plataforma aberta)",
        "descricao": "Antiga Odeprev, rebatizada Vexty em 2022 para ter identidade "
                     "própria (uma EFPC não pode ser vinculada ao nome de uma única "
                     "empresa). Fundada em 1995 pela então Odebrecht (hoje Novonor), "
                     "hoje atende mais de 110 empresas-patrocinadoras diferentes.",
        "fonte": "Confirmado - site institucional Vexty.",
    },
 
    4248: {  # VISÃO PREV
        "patrocinadoras": ["Grupo Telefônica | Vivo (diversas empresas do grupo no Brasil)"],
        "setor": "Telecomunicações",
        "descricao": "Criada em dezembro de 2004 para gerir os planos de previdência das "
                     "empresas do Grupo Telefônica/Vivo no Brasil. Ocupa a 24ª posição "
                     "entre as maiores EFPCs do país em volume de reservas, com mais de "
                     "21 mil participantes.",
        "fonte": "Confirmado - site institucional Visão Prev e Instituto Brasileiro de Atuária.",
    },
}

# Data da última atualização da base (para exibir no modal)
FICHAS_ATUALIZADO_EM = "2026-07-15"