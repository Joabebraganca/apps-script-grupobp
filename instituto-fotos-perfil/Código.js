/**
 * Padronização de fotos de perfil por domínio — Google Workspace (tenant Instituto)
 *
 * Mapeamento SEM PLANILHA: o nome do arquivo na pasta do Drive é o domínio.
 *   eonbr.com.png            -> todos os usuários @eonbr.com
 *   ironbergbrasilia.com.br.png -> todos os usuários @ironbergbrasilia.com.br
 *
 * Pré-requisitos:
 *  - Rodar com conta Superadministrador do tenant (sync@eonbr.com)
 *  - Serviços > adicionar "Admin SDK API" (identificador AdminDirectory, directory_v1)
 *  - Projeto Cloud vinculado com Admin SDK API habilitada
 *
 * Ordem de execução: listarFotos() -> testarUmUsuario() -> padronizarFotos() com
 * DRY_RUN true -> padronizarFotos() com DRY_RUN false.
 */

const CONFIG = {
  // Pasta "Fotos de perfil por domínio" — Drive da própria conta sync@eonbr.com
  PASTA_FOTOS_ID: '1ZJ1SC5Amy93G4ndyZ62L1Ev5dyG6RBHM',

  // true = só registra no log o que faria, sem alterar nada
  DRY_RUN: false,

  // false = não mexe em quem já tem foto (respeita quem já personalizou)
  SOBRESCREVER_FOTO_EXISTENTE: true,

  IGNORAR_PREFIXOS: ['no-reply', 'noreply', 'sync', 'admin-', 'teste'],
  IGNORAR_EMAILS: [],
  IGNORAR_OUS: ['/Automações'],

  // e-mail usado no testarUmUsuario()
  EMAIL_DE_TESTE: 'joabe.braganca@eonbr.com',

  // pausa e reagenda antes do limite de 6 min do Apps Script
  LIMITE_SEGUNDOS: 270
};

// ---------------------------------------------------------------- execução

/** 1) Confere o acesso da conta à pasta e mostra os nomes exatos dos arquivos. */
function listarFotos() {
  const arquivos = DriveApp.getFolderById(CONFIG.PASTA_FOTOS_ID).getFiles();
  const formatosOk = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff'];
  let n = 0;
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    const tipo = f.getBlob().getContentType();
    const alerta = formatosOk.indexOf(String(tipo).toLowerCase()) === -1 ? '  <-- FORMATO NÃO ACEITO' : '';
    Logger.log(f.getName() + '  |  ' + tipo + alerta);
    n++;
  }
  Logger.log('Total: ' + n + ' arquivo(s)');
}

/** 2) Testa em um único usuário antes de rodar no tenant inteiro. */
function testarUmUsuario() {
  const fotos = carregarFotosPorDominio_();
  const u = AdminDirectory.Users.get(CONFIG.EMAIL_DE_TESTE);
  const stats = { atualizados: 0, pulados: 0, semFoto: 0, erros: 0 };
  processarUsuario_(u, fotos, stats);
  Logger.log(JSON.stringify(stats));
}

/** 3) Função principal — roda no tenant inteiro. */
function padronizarFotos() {
  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  const fotos = carregarFotosPorDominio_();

  if (!Object.keys(fotos).length) {
    throw new Error('Nenhuma imagem encontrada na pasta ' + CONFIG.PASTA_FOTOS_ID);
  }
  Logger.log('Domínios com foto: ' + Object.keys(fotos).join(', '));

  let pageToken = props.getProperty('PAGE_TOKEN') || null;
  const stats = JSON.parse(props.getProperty('STATS') ||
    '{"atualizados":0,"pulados":0,"semFoto":0,"erros":0}');

  do {
    const resp = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 200,
      pageToken: pageToken || undefined,
      orderBy: 'email',
      projection: 'basic',
      query: 'isSuspended=false'
    });

    (resp.users || []).forEach(function (u) { processarUsuario_(u, fotos, stats); });
    pageToken = resp.nextPageToken || null;

    if (pageToken && (Date.now() - inicio) / 1000 > CONFIG.LIMITE_SEGUNDOS) {
      props.setProperty('PAGE_TOKEN', pageToken);
      props.setProperty('STATS', JSON.stringify(stats));
      agendarContinuacao_();
      Logger.log('Pausado por tempo — continuação agendada. Parcial: ' + JSON.stringify(stats));
      return;
    }
  } while (pageToken);

  props.deleteProperty('PAGE_TOKEN');
  props.deleteProperty('STATS');
  limparContinuacoes_();
  Logger.log((CONFIG.DRY_RUN ? '[SIMULAÇÃO] ' : '') + 'Concluído: ' + JSON.stringify(stats));
}

/** Alvo do gatilho de continuação — não chame direto. */
function continuarPadronizacao() {
  limparContinuacoes_();
  padronizarFotos();
}

/** Zera o progresso salvo (use se abortar no meio e quiser recomeçar do zero). */
function resetarProgresso() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('PAGE_TOKEN');
  props.deleteProperty('STATS');
  limparContinuacoes_();
  Logger.log('Progresso zerado.');
}

// ---------------------------------------------------------------- internos

function carregarFotosPorDominio_() {
  const arquivos = DriveApp.getFolderById(CONFIG.PASTA_FOTOS_ID).getFiles();
  const mapa = {};
  while (arquivos.hasNext()) {
    const f = arquivos.next();
    const nome = f.getName();
    const dominio = nome.replace(/\.(png|jpe?g|gif|bmp|tiff?)$/i, '').trim().toLowerCase();
    const blob = f.getBlob();
    mapa[dominio] = {
      photoData: Utilities.base64EncodeWebSafe(blob.getBytes()),
      mimeType: mimeType_(blob.getContentType()),
      arquivo: nome
    };
  }
  return mapa;
}

function mimeType_(contentType) {
  const t = String(contentType || '').toLowerCase();
  if (t.indexOf('png') !== -1) return 'PNG';
  if (t.indexOf('gif') !== -1) return 'GIF';
  if (t.indexOf('bmp') !== -1) return 'BMP';
  if (t.indexOf('tif') !== -1) return 'TIFF';
  return 'JPEG';
}

function processarUsuario_(u, fotos, stats) {
  const email = String(u.primaryEmail || '').toLowerCase();
  const local = email.split('@')[0];
  const dominio = email.split('@')[1];

  const ignorado =
    CONFIG.IGNORAR_EMAILS.some(function (e) { return e.toLowerCase() === email; }) ||
    CONFIG.IGNORAR_PREFIXOS.some(function (p) { return local.indexOf(p) === 0; }) ||
    CONFIG.IGNORAR_OUS.indexOf(u.orgUnitPath) !== -1;

  if (ignorado) { stats.pulados++; return; }

  const foto = fotos[dominio];
  if (!foto) {
    stats.semFoto++;
    Logger.log('Sem imagem para o domínio ' + dominio + ' (' + email + ')');
    return;
  }

  if (!CONFIG.SOBRESCREVER_FOTO_EXISTENTE && jaTemFoto_(email)) {
    stats.pulados++;
    return;
  }

  if (CONFIG.DRY_RUN) {
    Logger.log('[SIMULAÇÃO] ' + email + ' <- ' + foto.arquivo);
    stats.atualizados++;
    return;
  }

  try {
    comRetentativa_(function () {
      AdminDirectory.Users.Photos.update(
        { photoData: foto.photoData, mimeType: foto.mimeType }, email);
    });
    stats.atualizados++;
  } catch (e) {
    stats.erros++;
    Logger.log('ERRO em ' + email + ': ' + e.message);
  }
}

function jaTemFoto_(email) {
  try {
    const p = AdminDirectory.Users.Photos.get(email);
    return !!(p && p.photoData);
  } catch (e) {
    return false; // 404 = usuário sem foto
  }
}

function comRetentativa_(fn) {
  let espera = 1000;
  for (let i = 0; i < 5; i++) {
    try {
      return fn();
    } catch (e) {
      const m = String(e.message || '');
      const temporario = /rate|quota|backend|internal|503|500|429/i.test(m);
      if (!temporario || i === 4) throw e;
      Utilities.sleep(espera + Math.floor(Math.random() * 500));
      espera *= 2;
    }
  }
}

function agendarContinuacao_() {
  limparContinuacoes_();
  ScriptApp.newTrigger('continuarPadronizacao').timeBased().after(60 * 1000).create();
}

function limparContinuacoes_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'continuarPadronizacao') ScriptApp.deleteTrigger(t);
  });
}