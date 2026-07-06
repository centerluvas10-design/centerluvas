// ══════════ Center Luvas — Google Apps Script Backend ══════════
// Deploy as: Web App → Execute as Me → Anyone (anonymous)
// After each edit, deploy a NEW version (do NOT reuse the old deployment URL).

function doGet(e) {
  var params = e.parameter;
  var type   = params.type || '';
  var props  = PropertiesService.getScriptProperties();

  // ── Public config (adminPass is never returned here) ──
  if (type === 'config') {
    var raw = props.getProperty('cl_cfg');
    var cfg = raw ? JSON.parse(raw) : {};
    return jsonOut(JSON.stringify({
      pixKey:    cfg.pixKey    || '',
      pixName:   cfg.pixName   || '',
      pixCity:   cfg.pixCity   || '',
      whatsapp:  cfg.whatsapp  || '',
      lojaMsg:   cfg.lojaMsg   || '',
      lojaAtiva: cfg.lojaAtiva !== false
    }));
  }

  // ── Products ──
  if (type === 'produtos') {
    return jsonOut(props.getProperty('cl_produtos') || '[]');
  }

  // ── Product images ──
  if (type === 'prod_imgs') {
    var imgUrls, imgOld;
    try { imgUrls = JSON.parse(props.getProperty('cl_img_urls') || '{}'); } catch(e) { imgUrls = {}; }
    try { imgOld  = JSON.parse(props.getProperty('cl_imgs')     || '{}'); } catch(e) { imgOld  = {}; }
    // Merge: Drive URLs override old base64 thumbnails
    var merged = {};
    Object.keys(imgOld).forEach(function(k){ merged[k] = imgOld[k]; });
    Object.keys(imgUrls).forEach(function(k){ merged[k] = imgUrls[k]; });
    return jsonOut(JSON.stringify(merged));
  }

  // ── All orders (admin only) ──
  if (type === 'pedidos') {
    var adminPass = params.admin_pass || '';
    if (!checkAdminPass(adminPass, props)) return jsonOut(JSON.stringify({erro:'senha incorreta'}));
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Pedidos');
    if (!sheet) return jsonOut('[]');
    var rows  = sheet.getDataRange().getValues();
    var lista = [];
    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][0] || rows[i][0] === 'id') continue;
      lista.push({
        id:            String(rows[i][0]),
        data:          String(rows[i][1]),
        cliente_nome:  String(rows[i][2]),
        telefone:      String(rows[i][3]),
        endereco:      String(rows[i][4]),
        produtos:      String(rows[i][5]),
        total:         rows[i][6],
        pagamento:     String(rows[i][7]),
        obs:           String(rows[i][8]),
        cliente_email: String(rows[i][9]),
        status:        String(rows[i][10] || 'aguardando')
      });
    }
    return jsonOut(JSON.stringify(lista));
  }

  // ── Client's own orders ──
  if (type === 'meus_pedidos') {
    var token = params.token || '';
    if (!token) return jsonOut('[]');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clientEmail = '';
    var cSheet = ss.getSheetByName('CL_Clientes');
    if (cSheet) {
      var cRows = cSheet.getDataRange().getValues();
      for (var j = 0; j < cRows.length; j++) {
        if (String(cRows[j][6]) === token) { clientEmail = String(cRows[j][2]); break; }
      }
    }
    if (!clientEmail) return jsonOut('[]');
    var pSheet = ss.getSheetByName('CL_Pedidos');
    if (!pSheet) return jsonOut('[]');
    var pRows = pSheet.getDataRange().getValues();
    var meus  = [];
    for (var k = 0; k < pRows.length; k++) {
      if (!pRows[k][0] || pRows[k][0] === 'id') continue;
      if (String(pRows[k][9]).toLowerCase() === clientEmail.toLowerCase()) {
        meus.push({
          id:        String(pRows[k][0]),
          pedido_id: String(pRows[k][0]),
          data:      String(pRows[k][1]),
          produtos:  String(pRows[k][5]),
          total:     pRows[k][6],
          pagamento: String(pRows[k][7]),
          status:    String(pRows[k][10] || 'aguardando')
        });
      }
    }
    return jsonOut(JSON.stringify(meus));
  }

  // ── Clients list (admin only) ──
  if (type === 'clientes') {
    var adminPass = params.admin_pass || '';
    if (!checkAdminPass(adminPass, props)) return jsonOut(JSON.stringify({erro:'senha incorreta'}));
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Clientes');
    if (!sheet) return jsonOut('[]');
    var rows  = sheet.getDataRange().getValues();
    var lista = [];
    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][0] || rows[i][2] === 'email') continue; // skip empty rows and header
      lista.push({
        id:        String(rows[i][0]),
        nome:      String(rows[i][1]),
        email:     String(rows[i][2]),
        telefone:  String(rows[i][3]),
        criado_em: String(rows[i][5])
      });
    }
    return jsonOut(JSON.stringify(lista));
  }

  return jsonOut('{}');
}

function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var data;
  try { data = JSON.parse(e.postData.contents); }
  catch (ex) { return jsonOut(JSON.stringify({erro:'json invalido'})); }

  var tipo = data.tipo || '';

  // ── Admin login verification (server-side; tracks failures and emails password after 3 wrong attempts) ──
  if (tipo === 'admin_login') {
    var pass = data.password || '';

    // Load failure counter (reset automatically after 1 hour)
    var failsRaw = props.getProperty('cl_login_fails');
    var fails    = failsRaw ? JSON.parse(failsRaw) : {count:0, ts:0};
    if (Date.now() - fails.ts > 3600000) fails = {count:0, ts:0};

    if (checkAdminPass(pass, props)) {
      props.deleteProperty('cl_login_fails');
      return jsonOut(JSON.stringify({ok:true}));
    }

    // Wrong password
    fails.count += 1;
    fails.ts = Date.now();
    props.setProperty('cl_login_fails', JSON.stringify(fails));

    if (fails.count >= 3) {
      props.deleteProperty('cl_login_fails'); // reset so next cycle starts fresh
      var cfgRaw   = props.getProperty('cl_cfg');
      var adminCfg = cfgRaw ? JSON.parse(cfgRaw) : {};
      var adminPass = adminCfg.adminPass || '1234';
      var emailSent = false;
      var emailErro = '';
      try {
        MailApp.sendEmail({
          to:      'centerluvas10@gmail.com',
          subject: 'Center Luvas — Recuperação de senha do painel admin',
          body:    'Olá!\n\nApós 3 tentativas incorretas de login no painel administrativo, sua senha atual foi recuperada automaticamente:\n\n'
                 + 'Senha: ' + adminPass + '\n\n'
                 + 'Acesse o painel e altere a senha nas Configurações se necessário.\n\n'
                 + '— Center Luvas'
        });
        emailSent = true;
      } catch(ex) { emailErro = String(ex.message || ex); }
      return jsonOut(JSON.stringify({erro:'senha incorreta', email_enviado:emailSent, email_erro:emailErro}));
    }

    return jsonOut(JSON.stringify({erro:'senha incorreta', tentativas_restantes: 3 - fails.count}));
  }

  // ── Save config ──
  if (tipo === 'config') {
    var oldRaw = props.getProperty('cl_cfg');
    var old    = oldRaw ? JSON.parse(oldRaw) : {};
    var newCfg = {
      pixKey:    data.pixKey    !== undefined ? data.pixKey    : (old.pixKey    || ''),
      pixName:   data.pixName   !== undefined ? data.pixName   : (old.pixName   || ''),
      pixCity:   data.pixCity   !== undefined ? data.pixCity   : (old.pixCity   || ''),
      whatsapp:  data.whatsapp  !== undefined ? data.whatsapp  : (old.whatsapp  || ''),
      lojaMsg:   data.lojaMsg   !== undefined ? data.lojaMsg   : (old.lojaMsg   || ''),
      lojaAtiva: data.lojaAtiva !== undefined ? data.lojaAtiva : (old.lojaAtiva !== false),
      adminPass: data.adminPass || old.adminPass || '1234'
    };
    props.setProperty('cl_cfg', JSON.stringify(newCfg));
    return jsonOut(JSON.stringify({ok:true}));
  }

  // ── Save products ──
  if (tipo === 'produtos') {
    var prodStr = data.produtos; // already a JSON string (double-encoded by frontend)
    try { JSON.parse(prodStr); } catch(ex) { return jsonOut(JSON.stringify({erro:'produtos json invalido'})); }
    props.setProperty('cl_produtos', prodStr);
    return jsonOut(JSON.stringify({ok:true}));
  }

  // ── Upload product image to Google Drive ──
  if (tipo === 'produto_img_drive') {
    var prodId    = data.id  || '';
    var imgBase64 = data.img || '';
    if (!prodId || !imgBase64) return jsonOut(JSON.stringify({erro:'dados incompletos'}));
    var b64     = imgBase64.replace(/^data:image\/\w+;base64,/, '');
    var decoded = Utilities.base64Decode(b64);
    var blob    = Utilities.newBlob(decoded, 'image/jpeg', prodId + '.jpg');
    var folderIt = DriveApp.getFoldersByName('CL_Imagens');
    var folder   = folderIt.hasNext() ? folderIt.next() : DriveApp.createFolder('CL_Imagens');
    var imgIds; try { imgIds = JSON.parse(props.getProperty('cl_img_ids') || '{}'); } catch(e) { imgIds = {}; }
    if (imgIds[prodId]) {
      try { DriveApp.getFileById(imgIds[prodId]).setTrashed(true); } catch(ex) {}
    }
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var url    = 'https://drive.google.com/uc?export=view&id=' + fileId;
    var imgUrls; try { imgUrls = JSON.parse(props.getProperty('cl_img_urls') || '{}'); } catch(e) { imgUrls = {}; }
    imgUrls[prodId] = url;
    props.setProperty('cl_img_urls', JSON.stringify(imgUrls));
    imgIds[prodId] = fileId;
    props.setProperty('cl_img_ids', JSON.stringify(imgIds));
    return jsonOut(JSON.stringify({ok:true, url:url}));
  }

  // ── Delete product image ──
  if (tipo === 'delete_img') {
    var prodId = data.id || '';
    var imgs; try { imgs = JSON.parse(props.getProperty('cl_imgs') || '{}'); } catch(e) { imgs = {}; }
    delete imgs[prodId];
    props.setProperty('cl_imgs', JSON.stringify(imgs));
    var imgUrls; try { imgUrls = JSON.parse(props.getProperty('cl_img_urls') || '{}'); } catch(e) { imgUrls = {}; }
    delete imgUrls[prodId];
    props.setProperty('cl_img_urls', JSON.stringify(imgUrls));
    var imgIds; try { imgIds = JSON.parse(props.getProperty('cl_img_ids') || '{}'); } catch(e) { imgIds = {}; }
    if (imgIds[prodId]) {
      try { DriveApp.getFileById(imgIds[prodId]).setTrashed(true); } catch(ex) {}
    }
    delete imgIds[prodId];
    props.setProperty('cl_img_ids', JSON.stringify(imgIds));
    return jsonOut(JSON.stringify({ok:true}));
  }

  // ── Register new client account ──
  if (tipo === 'registro') {
    var email     = (data.email || '').toLowerCase().trim();
    var nome      = data.nome || '';
    var telefone  = data.telefone || '';
    var senhaHash = data.senha_hash || '';
    if (!email || !nome || !senhaHash) return jsonOut(JSON.stringify({erro:'dados incompletos'}));
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Clientes');
    if (!sheet) {
      sheet = ss.insertSheet('CL_Clientes');
      sheet.appendRow(['id','nome','email','telefone','senha_hash','criado_em','token']);
    }
    var rows = sheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][2]).toLowerCase() === email) return jsonOut(JSON.stringify({erro:'e-mail já cadastrado'}));
    }
    var id    = 'cli_' + Date.now();
    var token = Utilities.getUuid();
    sheet.appendRow([id, nome, email, telefone, senhaHash, new Date().toISOString(), token]);
    return jsonOut(JSON.stringify({ok:true, token:token, nome:nome}));
  }

  // ── Client login ──
  if (tipo === 'login') {
    var email     = (data.email || '').toLowerCase().trim();
    var senhaHash = data.senha_hash || '';
    if (!email || !senhaHash) return jsonOut(JSON.stringify({erro:'dados incompletos'}));
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Clientes');
    if (!sheet) return jsonOut(JSON.stringify({erro:'usuário não encontrado'}));
    var rows = sheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][2]).toLowerCase() === email) {
        var storedHash = String(rows[i][4]);
        var isTemp     = storedHash.indexOf('TEMP:') === 0;
        var realHash   = isTemp ? storedHash.slice(5) : storedHash;
        if (realHash === senhaHash) {
          var token = Utilities.getUuid();
          sheet.getRange(i + 1, 7).setValue(token);
          var resp = {ok:true, token:token, nome:String(rows[i][1])};
          if (isTemp) resp.needs_password_change = true;
          return jsonOut(JSON.stringify(resp));
        }
        return jsonOut(JSON.stringify({erro:'senha incorreta'}));
      }
    }
    return jsonOut(JSON.stringify({erro:'usuário não encontrado'}));
  }

  // ── Change password after temp login ──
  if (tipo === 'trocar_senha') {
    var token        = data.token || '';
    var novaSenhaHash = data.nova_senha_hash || '';
    if (!token || !novaSenhaHash) return jsonOut(JSON.stringify({erro:'dados incompletos'}));
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Clientes');
    if (!sheet) return jsonOut(JSON.stringify({erro:'usuário não encontrado'}));
    var rows = sheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][6]) === token) {
        sheet.getRange(i + 1, 5).setValue(novaSenhaHash); // store plain hash, no TEMP: prefix
        return jsonOut(JSON.stringify({ok:true}));
      }
    }
    return jsonOut(JSON.stringify({erro:'sessão inválida'}));
  }

  // ── New store order ──
  if (tipo === 'pedido') {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Pedidos');
    if (!sheet) {
      sheet = ss.insertSheet('CL_Pedidos');
      sheet.appendRow(['id','data','cliente_nome','telefone','endereco','produtos','total','pagamento','obs','cliente_email','status']);
    }
    // Frontend sends: {tipo, id, data, cliente (string name), telefone, endereco, produtos (string), total, pagamento, obs, cliente_email}
    sheet.appendRow([
      data.id            || ('PED' + Date.now()),
      data.data          || new Date().toISOString(),
      data.cliente       || data.cliente_nome || '',
      data.telefone      || '',
      data.endereco      || '',
      data.produtos      || '',
      parseFloat(data.total) || 0,
      data.pagamento     || '',
      data.obs           || '',
      data.cliente_email || '',
      data.status        || 'aguardando'
    ]);
    return jsonOut(JSON.stringify({ok:true}));
  }

  // ── Update order status ──
  if (tipo === 'atualizar_status') {
    var pedidoId = String(data.pedido_id || '');
    var status   = String(data.status || '');
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Pedidos');
    if (!sheet) return jsonOut(JSON.stringify({erro:'planilha não encontrada'}));
    var rows = sheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === pedidoId) {
        sheet.getRange(i + 1, 11).setValue(status); // column K = status
        return jsonOut(JSON.stringify({ok:true}));
      }
    }
    return jsonOut(JSON.stringify({erro:'pedido não encontrado'}));
  }

  // ── Admin manual sale (stock already updated via syncProdutos) ──
  if (tipo === 'registrar_venda') {
    return jsonOut(JSON.stringify({ok:true}));
  }

  // ── Register expense ──
  if (tipo === 'registrar_despesa') {
    return jsonOut(JSON.stringify({ok:true}));
  }

  // ── Reset a client's password (admin action) ──
  if (tipo === 'reset_senha') {
    if (!checkAdminPass(data.admin_pass || '', props)) return jsonOut(JSON.stringify({erro:'senha incorreta'}));
    var email = (data.email || '').toLowerCase().trim();
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('CL_Clientes');
    if (!sheet) return jsonOut(JSON.stringify({erro:'cliente não encontrado'}));
    var rows = sheet.getDataRange().getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][2]).toLowerCase() === email) {
        var tempSenha = Math.random().toString(36).substr(2, 8);
        var digest    = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, tempSenha, Utilities.Charset.UTF_8);
        var hexHash   = digest.map(function(b){ return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
        // Store with TEMP: prefix so login detects it and forces password change
        sheet.getRange(i + 1, 5).setValue('TEMP:' + hexHash);
        return jsonOut(JSON.stringify({ok:true, temp_senha:tempSenha}));
      }
    }
    return jsonOut(JSON.stringify({erro:'cliente não encontrado'}));
  }

  return jsonOut(JSON.stringify({erro:'tipo desconhecido: ' + tipo}));
}

// ── Helpers ──

// ── Run this function ONCE to clear corrupted image data from PropertiesService ──
function limparImgCache() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('cl_imgs');
  Logger.log('cl_imgs apagado com sucesso.');
}

// ── Run this function ONCE from the Apps Script editor to authorize DriveApp ──
function testarDrive() {
  var folder = DriveApp.createFolder('CL_Imagens_Teste');
  folder.setTrashed(true);
  Logger.log('DriveApp autorizado com sucesso.');
}

// ── Run this function ONCE from the Apps Script editor to authorize Gmail ──
function testarEmail() {
  MailApp.sendEmail({
    to:      'centerluvas10@gmail.com',
    subject: 'Center Luvas — Teste de e-mail',
    body:    'E-mail de teste enviado com sucesso. O sistema de recuperação de senha está autorizado.'
  });
  Logger.log('E-mail enviado com sucesso.');
}

function checkAdminPass(pass, props) {
  var raw    = props.getProperty('cl_cfg');
  var cfg    = raw ? JSON.parse(raw) : {};
  var stored = cfg.adminPass || '1234';
  return pass === stored;
}

function jsonOut(str) {
  return ContentService.createTextOutput(str)
    .setMimeType(ContentService.MimeType.TEXT);
}
