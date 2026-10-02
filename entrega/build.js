const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/be72551b-f99c-4a63-8538-7dcec9ac7c60_0bc1472f-8aea-44c1-b595-625272365ae1/pptx/scripts/apply_theme.js");
const A = (f) => __dirname + "/assets/" + f;
// Paleta — Manual de Marca Thunders, p.13
const NAVY="001845", AZUL="002872", ROYAL="03399D", CYAN="09C6FF", TEAL="00A0D1", BG="F3F7FE", WHITE="FFFFFF";
const HEAD="DM Sans", BODY="DM Sans"; // Manual p.20 (Elza, p.21, trocada por DM Sans: glifos quebrados na máquina do usuário)
const THEME={name:"Thunders",headFontFace:HEAD,bodyFontFace:BODY,colors:{dk1:NAVY,lt1:WHITE,dk2:AZUL,lt2:BG,accent1:ROYAL,accent2:TEAL,accent3:CYAN,accent4:AZUL,accent5:NAVY,accent6:BG,hlink:ROYAL,folHlink:AZUL}};

const pres = new pptxgen(); pres.layout="LAYOUT_WIDE"; pres.theme={headFontFace:HEAD,bodyFontFace:BODY};
pres.title="Processo de Desenvolvimento de Software — Thunders"; pres.author="Thunders";
const R=0.14; // cantos arredondados (~20px proporcional) — Manual p.26

pres.defineSlideMaster({title:"CAPA",background:{path:A("cover_bg.png")},objects:[]});
pres.defineSlideMaster({title:"CONTEUDO",background:{color:BG},
  objects:[{image:{path:A("logo_reduzida_light.png"),x:0.6,y:6.88,w:1.14,h:0.34}}],
  slideNumber:{x:12.2,y:6.9,w:0.55,h:0.3,fontFace:BODY,fontSize:11,color:AZUL,align:"left"}});

const T=(s,text,o)=>s.addText(text,Object.assign({isTextBox:true,fontFace:BODY,color:NAVY,margin:0,valign:"top",align:"left"},o));
function head(s,label,title){
  T(s,label,{x:0.6,y:0.45,w:8,h:0.3,fontFace:HEAD,fontSize:12,bold:true,color:ROYAL,charSpacing:2});
  T(s,title,{x:0.6,y:0.8,w:12.1,h:1.0,fontFace:HEAD,fontSize:30,bold:true,color:NAVY,valign:"top"});
}
function card(s,x,y,w,h,fill,name){s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,fill:{color:fill||WHITE},line:{color:fill||WHITE,width:0},rectRadius:R,objectName:name||"card",shadow:fill?undefined:{type:"outer",color:"001845",opacity:0.10,blur:8,offset:2,angle:90}});}
function chip(s,text,x,y,w,fill,color,fs){card(s,x,y,w,0.38,fill);T(s,text,{x:x+0.12,y,w:w-0.24,h:0.38,fontSize:fs||11,color,valign:"middle",bold:false});}
function bullets(s,items,o){T(s,items.map((t,i)=>({text:t,options:{bullet:{indent:14},breakLine:i<items.length-1,paraSpaceAfter:5}})),o);}
const notes=(s,t)=>s.addNotes(t);

// 1 CAPA
{const s=pres.addSlide({masterName:"CAPA"});
 s.addImage({path:A("raio_outline.png"),x:8.7,y:0.7,w:3.9,h:4.79,objectName:"raio"});
 s.addImage({path:A("logo_reduzida_dark.png"),x:0.8,y:0.7,w:1.9,h:0.56});
 T(s,"Processo de Desenvolvimento de Software",{x:0.8,y:2.3,w:8,h:1.9,fontFace:HEAD,fontSize:40,bold:true,color:WHITE});
 T(s,"Da passagem do briefing à entrega da garantia",{x:0.8,y:4.75,w:7.6,h:0.6,fontFace:HEAD,fontSize:20,color:CYAN});
 T(s,"Resumo do quadro de processo — Miro",{x:0.8,y:6.7,w:7,h:0.3,fontSize:12,color:WHITE});
 notes(s,"Fonte: quadro Miro (frame 'Processo - Em Aprimoramento'). Layout e cores: Manual de Marca Thunders (p.13, p.17, p.25).");}

// 2 VISÃO GERAL
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"VISÃO GERAL","Três perguntas guiam a célula da imersão à entrega");
 const q=[["O que devemos construir?","Concepção da Solução","Dores, oportunidades, hipóteses e premissas viram o Plano da Iniciativa."],
  ["Devemos executar?","Gate","Coerência com as dores, alinhamento estratégico e viabilidade técnica."],
  ["Como devemos implementar a feature?","Build","Define, implementa e testa até a entrega."]];
 q.forEach((c,i)=>{const x=0.6+i*4.1; card(s,x,2.0,3.9,2.35,WHITE);
  T(s,c[1].toUpperCase(),{x:x+0.3,y:2.2,w:3.3,h:0.3,fontFace:HEAD,fontSize:12,bold:true,color:TEAL,charSpacing:2});
  T(s,c[0],{x:x+0.3,y:2.55,w:3.3,h:0.9,fontFace:HEAD,fontSize:20,bold:true,color:NAVY});
  T(s,c[2],{x:x+0.3,y:3.5,w:3.3,h:0.75,fontSize:13,color:AZUL});});
 T(s,"Hierarquia do trabalho",{x:0.6,y:4.7,w:5,h:0.3,fontFace:HEAD,fontSize:14,bold:true,color:NAVY});
 ["Iniciativa","Feature","User Stories","Tasks (código)"].forEach((t,i)=>{chip(s,t,0.6+i*1.5,5.1,1.4,ROYAL,WHITE,11);});
 card(s,6.9,4.7,5.8,1.85,AZUL);
 T(s,"Célula autônoma",{x:7.2,y:4.88,w:5.2,h:0.3,fontFace:HEAD,fontSize:14,bold:true,color:CYAN});
 T(s,"Propõe melhorias na própria dinâmica (retrospectivas técnicas) e no processo, nos agentes, nas skills e na arquitetura, com os respectivos responsáveis.",{x:7.2,y:5.25,w:5.2,h:1.2,fontSize:13,color:WHITE});
 T(s,"A gestão guia a célula da imersão até a conclusão da iniciativa.",{x:0.6,y:5.75,w:5.9,h:0.6,fontSize:13,color:AZUL});
 notes(s,"Fonte: Miro, frame 'Processo - Em Aprimoramento' — perguntas 'O que devemos construir?', 'Devemos executar?', 'Como devemos implementar a feature?'; blocos 'Iniciativa/Feature/User Stories/Tasks'; 'Gestão'; 'Detalhes importantes' (célula autônoma). O quadro não traz um objetivo escrito do processo — ⚠️ a confirmar.");}

// 3 ESPINHA
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"PROCESSO MACRO","Espinha geral do processo");
 const st=[["Passagem do Briefing","A célula recebe o problema a ser tratado"],
  ["Discovery","Imersão e Concepção da Solução"],
  ["Gate","Devemos executar?"],
  ["Build","Define, implementa e testa"],
  ["Release","Publicação, divulgada 24 h antes"],
  ["GMUD","Itens indispensáveis e checklist de subida"],
  ["Pós go-live","Monitoramento de uso e performance"],
  ["Entrega da Garantia","Evidências de uso, performance e suporte"]];
 const w=1.45,gap=0.0,x0=0.6,pitch=(12.13-w)/7, ly=2.55;
 s.addShape(pres.shapes.LINE,{x:x0+w/2,y:ly,w:pitch*7,h:0,line:{color:ROYAL,width:3,endArrowType:"triangle"},objectName:"linha"});
 st.forEach((e,i)=>{const x=x0+i*pitch, cx=x+w/2;
  const col=i==2?AZUL:(i<4?ROYAL:(i<6?TEAL:NAVY));
  if(i==2) s.addShape(pres.shapes.DIAMOND,{x:cx-0.32,y:ly-0.32,w:0.64,h:0.64,fill:{color:CYAN},line:{color:CYAN,width:0}});
  else s.addShape(pres.shapes.OVAL,{x:cx-0.27,y:ly-0.27,w:0.54,h:0.54,fill:{color:col},line:{color:BG,width:3}});
  T(s,String(i+1),{x:cx-0.27,y:ly-0.27,w:0.54,h:0.54,fontFace:HEAD,fontSize:14,bold:true,color:i==2?NAVY:WHITE,align:"center",valign:"middle"});
  card(s,x,3.25,w,2.55,WHITE);
  T(s,e[0],{x:x+0.12,y:3.4,w:w-0.24,h:0.85,fontFace:HEAD,fontSize:14,bold:true,color:NAVY});
  T(s,e[1],{x:x+0.12,y:4.3,w:w-0.24,h:1.4,fontSize:12,color:AZUL});});
 T(s,"Gate reprovado volta ao passo anterior (ou replaneja / abandona). Teste não aprovado gera ajustes dentro do tempo de Build.",{x:0.6,y:6.05,w:12,h:0.5,fontSize:13,color:AZUL});
 notes(s,"Espinha fornecida por Sidney, confirmada nas respostas às divergências: Gate único antes do Build; antes da publicação não há gate (há divulgação de Release); depois segue para o monitoramento. Release = Publicação; Discovery = Concepção da Solução. Conteúdo de cada etapa: Miro, frame 'Processo - Em Aprimoramento'. ⚠️ Ordem Release → GMUD inferida do quadro (blocos verdes) — a confirmar.");}

// 4 BRIEFING + IMERSÃO
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"1 · PASSAGEM DO BRIEFING E IMERSÃO","A célula recebe o tema e entende o suficiente");
 card(s,0.6,2.0,5.4,3.5,WHITE);
 T(s,"Passagem do Briefing",{x:0.9,y:2.2,w:4.8,h:0.4,fontFace:HEAD,fontSize:20,bold:true,color:NAVY});
 bullets(s,["O tema que iremos trabalhar","A célula recebe o problema a ser tratado","Cerimônia: Passagem do Briefing"],{x:0.9,y:2.8,w:4.8,h:1.9,fontSize:16,color:AZUL});
 s.addShape(pres.shapes.RIGHT_ARROW,{x:6.15,y:3.45,w:0.7,h:0.5,fill:{color:TEAL},line:{color:TEAL,width:0}});
 card(s,7.0,2.0,5.7,3.5,WHITE);
 T(s,"Imersão",{x:7.3,y:2.2,w:5.1,h:0.4,fontFace:HEAD,fontSize:20,bold:true,color:NAVY});
 bullets(s,["Célula entendendo o tema","Entender o suficiente para poder trabalhar no tema","Primeiras dores, oportunidades, hipóteses e premissas são levantadas","A célula precisa entender um pouco de tudo do que está sendo elaborado"],{x:7.3,y:2.8,w:5.1,h:2.6,fontSize:16,color:AZUL});
 card(s,0.6,5.75,12.1,0.8,AZUL);
 T(s,[{text:"Indicador:  ",options:{bold:true,color:CYAN}},{text:"Tempo da Imersão",options:{color:WHITE}}],{x:0.9,y:5.75,w:11.5,h:0.8,fontSize:15,valign:"middle"});
 notes(s,"Fonte: Miro, frame 'Processo - Em Aprimoramento' — blocos 'Passagem do Briefing' ('o tema que iremos trabalhar', 'A célula recebe o problema a ser tratado', cerimônia) e 'Imersão' ('Célula entendendo o tema', 'Entender o suficiente…', indicador 'Tempo da Imersão').");}

// 5 DISCOVERY + GATE
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"2 · DISCOVERY","Concepção da Solução e Gate");
 card(s,0.6,2.0,7.0,4.55,WHITE);
 T(s,"Concepção da Solução",{x:0.9,y:2.15,w:6,h:0.4,fontFace:HEAD,fontSize:20,bold:true,color:NAVY});
 T(s,"O que devemos construir?",{x:0.9,y:2.6,w:6,h:0.3,fontSize:14,color:TEAL,bold:true});
 chip(s,"Dores / Oportunidades",0.9,3.05,2.9,ROYAL,WHITE,12); chip(s,"Hipóteses / Premissas",4.0,3.05,2.9,ROYAL,WHITE,12);
 T(s,"Validar, questionar e entender com o mundo externo; discussão e validação totalmente em conjunto.",{x:0.9,y:3.6,w:6.4,h:0.7,fontSize:13,color:AZUL});
 T(s,"Plano da Iniciativa",{x:0.9,y:4.4,w:6,h:0.3,fontFace:HEAD,fontSize:14,bold:true,color:NAVY});
 ["Negócio","Solução","Solução arquitetural","Governança"].forEach((t,i)=>chip(s,t,0.9+(i%2)*3.1,4.8+Math.floor(i/2)*0.48,2.9,BG,NAVY,12));
 T(s,[{text:"Output: ",options:{bold:true}},{text:"Plano da Iniciativa e Roadmap priorizado  ·  Indicador: Tempo em Concepção"}],{x:0.9,y:5.85,w:6.4,h:0.6,fontSize:12,color:AZUL});
 card(s,7.9,2.0,4.8,4.55,AZUL);
 s.addShape(pres.shapes.DIAMOND,{x:8.2,y:2.2,w:0.5,h:0.5,fill:{color:CYAN},line:{color:CYAN,width:0}});
 T(s,"Gate",{x:8.85,y:2.15,w:3.5,h:0.4,fontFace:HEAD,fontSize:20,bold:true,color:WHITE});
 T(s,"Devemos executar?",{x:8.85,y:2.55,w:3.5,h:0.3,fontSize:13,color:CYAN,bold:true});
 bullets(s,["Solução coerente com as dores e os temas","Alinhada com a estratégia da empresa","Solução técnica viável (tecnologia, custo)"],{x:8.2,y:3.1,w:4.3,h:1.4,fontSize:13,color:WHITE});
 T(s,[{text:"Aprovado: ",options:{bold:true,color:CYAN}},{text:"Build",options:{breakLine:true}},{text:"Reprovado: ",options:{bold:true,color:CYAN}},{text:"volta ao passo anterior",options:{breakLine:true}},{text:"Replanejamento: ",options:{bold:true,color:CYAN}},{text:"a célula não pode seguir",options:{breakLine:true}},{text:"Abandono: ",options:{bold:true,color:CYAN}},{text:"custo ou know-how"}].map(r=>Object.assign(r,{options:Object.assign({color:WHITE,paraSpaceAfter:3},r.options)})),{x:8.2,y:4.55,w:4.3,h:1.35,fontSize:12});
 T(s,"Indicadores: taxa de reprovação, iniciativas replanejadas e abandonadas",{x:8.2,y:5.95,w:4.3,h:0.5,fontSize:11,color:WHITE});
 notes(s,"Fonte: Miro, frame 'Processo - Em Aprimoramento' — 'Concepção da Solução' (Dores/Oportunidades, Hipóteses/Premissas, Plano da Iniciativa, Roadmap priorizado, Tempo em Concepção) e 'Gate / Abandono da Iniciativa' (perguntas, Aprovado/Reprovado/Replanejamento/Abandono, indicadores, cerimônia Avaliação do Gate). Observação: a numeração do Plano da Iniciativa no quadro pula o item 3.");}

// 6 BUILD
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"3 · BUILD","Define, implementa e testa");
 const f=[["Define","Como devemos implementar a feature?\nSDD · Protótipo"],["Implementa","Tasks (código) por feature"],["Teste","Funcional e usabilidade"]];
 f.forEach((c,i)=>{const x=0.6+i*3.3; card(s,x,2.0,2.9,2.05,WHITE);
  T(s,String(i+1),{x:x+0.25,y:2.15,w:0.5,h:0.5,fontFace:HEAD,fontSize:26,bold:true,color:TEAL});
  T(s,c[0],{x:x+0.25,y:2.7,w:2.4,h:0.4,fontFace:HEAD,fontSize:18,bold:true,color:NAVY});
  T(s,c[1],{x:x+0.25,y:3.15,w:2.45,h:0.85,fontSize:12,color:AZUL});
  if(i<2) s.addShape(pres.shapes.RIGHT_ARROW,{x:x+2.95,y:2.85,w:0.3,h:0.35,fill:{color:TEAL},line:{color:TEAL,width:0}});});
 s.addShape(pres.shapes.DIAMOND,{x:10.95,y:2.1,w:1.1,h:1.1,fill:{color:CYAN},line:{color:CYAN,width:0}});
 T(s,"OK?",{x:10.95,y:2.1,w:1.1,h:1.1,fontFace:HEAD,fontSize:16,bold:true,color:NAVY,align:"center",valign:"middle"});
 T(s,"Sim: sem bugs nem ajustes — segue para o Release",{x:10.35,y:3.3,w:2.35,h:0.75,fontSize:12,color:AZUL});
 card(s,0.6,4.3,12.1,0.75,AZUL);
 T(s,[{text:"Não: ",options:{bold:true,color:CYAN}},{text:"Identificação de ajustes — a célula ajusta e isso entra no tempo de Build.",options:{color:WHITE}}],{x:0.9,y:4.3,w:11.5,h:0.75,fontSize:14,valign:"middle"});
 T(s,"Indicadores de fluxo",{x:0.6,y:5.3,w:5,h:0.3,fontFace:HEAD,fontSize:14,bold:true,color:NAVY});
 ["WIP","Idade do WIP","Lead Time","Cycle Time","Tempo em Definição","Tempo em Teste","Tempo em Homologação*","Bug em Homologação*"].forEach((t,i)=>chip(s,t,0.6+(i%4)*3.05,5.7+Math.floor(i/4)*0.48,2.9,WHITE,NAVY,11));
 notes(s,"Fonte: Miro, frame 'Processo - Em Aprimoramento' — bloco 'Build' (Roadmap priorizado, 'Como devemos implementar a feature?', SDD, Protótipo, Implementa/Tasks, Teste funcional e usabilidade, 'OK?', 'Identificação de Ajustes', indicadores de fluxo). * Indicadores com asterisco no quadro.");}

// 7 RELEASE + GMUD
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"4 · RELEASE E GMUD","Release divulgada antes da subida, com GMUD");
 card(s,0.6,2.0,4.3,4.55,WHITE);
 T(s,"Release (assíncrona)",{x:0.9,y:2.15,w:3.8,h:0.4,fontFace:HEAD,fontSize:18,bold:true,color:NAVY});
 bullets(s,["Divulgada no mínimo 24 h antes da subida","Postada no canal do Teams para todos","Breve descrição, Demos e Documento Técnico (Requisitos de Subida)","A célula decide qual pacote de versões sobe (precisa gerar valor)"],{x:0.9,y:2.75,w:3.8,h:2.7,fontSize:13,color:AZUL});
 T(s,"Indicadores: releases abortadas e concluídas com sucesso, ocorrências em produção, tempo de execução",{x:0.9,y:5.5,w:3.8,h:0.9,fontSize:11,color:AZUL});
 const m=[["Negócio","Publica no SharePoint, com informações no Jira: features, regra de negócio, vídeo de utilização"],
  ["Acompanhamento / utilização","Features, script de acompanhamento de uso (quando aplicável), TAG(s) do Mixpanel"],
  ["Checklist de Subida","Features, tamanho de campos, índices, slow queries e plano de execução"],
  ["GMUD: itens indispensáveis","Microsserviços (PRs), feature flag, tenants e scripts; informações centralizadas na feature do Jira"]];
 m.forEach((c,i)=>{const x=5.2+(i%2)*3.8,y=2.0+Math.floor(i/2)*2.35; card(s,x,y,3.65,2.2,i==3?AZUL:WHITE);
  T(s,c[0],{x:x+0.25,y:y+0.18,w:3.2,h:0.55,fontFace:HEAD,fontSize:15,bold:true,color:i==3?CYAN:ROYAL});
  T(s,c[1],{x:x+0.25,y:y+0.8,w:3.2,h:1.3,fontSize:12,color:i==3?WHITE:AZUL});});
 notes(s,"Fonte: Miro, frame 'Processo - Em Aprimoramento' — bloco 'Deploy / Release' (cerimônia Release assíncrona, Teams, 24 h, Versão) e blocos verdes 'Release': Negócio, Acompanhamento/utilização, Checklist de Subida, 'GMUD - Itens indispensáveis' e nota 'Processo de GMUD'. Segundo Sidney, cada bloco verde é um micro-processo e as notas são pontos importantes. ⚠️ Confirmar se 'duas etapas' de GMUD = Checklist de Subida + GMUD itens indispensáveis.");}

// 8 PÓS GO-LIVE + GARANTIA
{const s=pres.addSlide({masterName:"CONTEUDO"});
 head(s,"5 · PÓS GO-LIVE E GARANTIA","Monitoramento e entrega da garantia");
 card(s,0.6,2.0,5.9,2.75,WHITE);
 T(s,"Pós go-live: Monitoramento",{x:0.9,y:2.15,w:5.3,h:0.4,fontFace:HEAD,fontSize:18,bold:true,color:NAVY});
 bullets(s,["Métricas de uso e de performance (Grafana)","Suporte (tickets) e fila do Rabbit (processamento, erro, DLQ)","Comparações: hoje × ontem, × última semana, × último mês (relativo)"],{x:0.9,y:2.75,w:5.3,h:1.9,fontSize:15,color:AZUL});
 card(s,6.8,2.0,5.9,2.75,WHITE);
 T(s,"Entrega da Garantia: Evidências",{x:7.1,y:2.15,w:5.3,h:0.4,fontFace:HEAD,fontSize:18,bold:true,color:NAVY});
 bullets(s,["Evidências de uso e de performance (Grafana)","Evidências de suporte (tickets) e da fila do Rabbit","Garantia de 90 dias, sempre sem custo ao cliente"],{x:7.1,y:2.75,w:5.3,h:1.9,fontSize:15,color:AZUL});
 T(s,"Indicadores de resultado",{x:0.6,y:5.1,w:6,h:0.3,fontFace:HEAD,fontSize:14,bold:true,color:NAVY});
 ["Deployment Frequency (mensal)","Change Failure Rate","Tempo gasto nas correções","Previsto × Real (utilização mensal)","Chamados em produção","Aderência em produção"].forEach((t,i)=>chip(s,t,0.6+(i%3)*4.08,5.5+Math.floor(i/3)*0.5,3.95,AZUL,WHITE,12));
 notes(s,"Fonte: Miro, frame 'Processo - Em Aprimoramento' — blocos verdes 'Pós go-live' (Monitoramento: métricas de uso, Grafana, suporte, Rabbit, comparações), 'Entrega da Garantia' (Evidências) e 'Indicadores de Resultado'. '90 dias, sem custo' vem do frame 'Thunders - Fluxo do Processo de Execução do Projeto' (6 etapas: ... 5. Garantia - 90 dias; 6. Sustentação) — ⚠️ confirmar se se aplica a este processo de desenvolvimento.");}

(async()=>{await pres.writeFile({fileName:__dirname+"/Processo_Desenvolvimento_Thunders.pptx"});await applyTheme(__dirname+"/Processo_Desenvolvimento_Thunders.pptx",THEME);console.log("ok");})();
