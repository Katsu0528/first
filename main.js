function AdCore() {
  try {
    Logger.log("実行開始");
    const base = SpreadsheetApp.getActiveSpreadsheet();

    // シート取得＆nullチェック
    const inputSheet = base.getSheetByName("入稿用");
    if (!inputSheet) throw new Error("「入稿用」シートが見つかりません。");
    const masterSheet = base.getSheetByName("マスタ");
    if (!masterSheet) throw new Error("「マスタ」シートが見つかりません。");
    const materialSheet = base.getSheetByName("素材追加");
    if (!materialSheet) throw new Error("「素材追加」シートが見つかりません。");
    const varMasterSheet = base.getSheetByName("変数マスタ");
    if (!varMasterSheet) throw new Error("「変数マスタ」シートが見つかりません。");

    const registerAd = String(inputSheet.getRange("B4").getValue()).toLowerCase() === 'true';
    const onlyAd = String(inputSheet.getRange("B5").getValue()).toLowerCase() === 'true';
    const onlyMaterial = String(inputSheet.getRange("B6").getValue()).toLowerCase() === 'true';

    // モード名の判別
    let modeName = "";
    if (registerAd) modeName = "広告登録＋素材追加";
    else if (onlyAd) modeName = "広告登録のみ";
    else if (onlyMaterial) modeName = "素材追加のみ";

    const modeFlags = [registerAd, onlyAd, onlyMaterial].filter(flag => flag);
    if (modeFlags.length !== 1) {
      SpreadsheetApp.getUi().alert(
        "エラー",
        "登録モード（B4:広告登録＋素材追加, B5:広告登録のみ, B6:素材追加のみ）のいずれか1つのみをチェックしてください。",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      throw new Error("登録モード（B4:広告登録＋素材追加, B5:広告登録のみ, B6:素材追加のみ）のいずれか1つのみをチェックしてください。");
    }

    // モード確認のポップアップ
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      "処理モード確認",
      "選択されたモードは【" + modeName + "】です。\nよろしいですか？",
      ui.ButtonSet.YES_NO
    );
    if (result === ui.Button.NO) {
      Logger.log("ユーザーが「いいえ」を選択したため処理を中断しました。");
      return;
    }

    if (registerAd) {
      const result = registerAdFromSheet();
      Logger.log("registerAdFromSheetの戻り値: " + JSON.stringify(result));
      if (!result || !result.promotionId) {
        SpreadsheetApp.getUi().alert("エラー", "広告登録結果が取得できませんでした", SpreadsheetApp.getUi().ButtonSet.OK);
        throw new Error("広告登録結果が取得できませんでした");
      }
      // 提携申請
      Logger.log("registerPromotionApply呼び出し: promotionId=" + result.promotionId);
      registerPromotionApply(result.promotionId);
      Logger.log("addMaterialFromSheet呼び出し: promotionName=" + result.promotionName + ", promotionId=" + result.promotionId);
      addMaterialFromSheet(result.promotionName, result.promotionId);
    } else if (onlyMaterial) {
      Logger.log("addMaterialFromSheet呼び出し（素材追加のみモード）");
      addMaterialFromSheet();
    } else if (onlyAd) {
      const result = registerAdFromSheet();
      // 提携申請
      if (result && result.promotionId) {
        registerPromotionApply(result.promotionId);
      }
    }
    Logger.log("実行完了");
  } catch (e) {
    Logger.log("エラー: " + e);
    SpreadsheetApp.getUi().alert("エラー", String(e), SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}
function registerAdFromSheet() { 
  const base = SpreadsheetApp.getActiveSpreadsheet();
  const entrySheet = base.getSheetByName("広告登録");
  if (!entrySheet) throw new Error("「広告登録」シートが見つかりません。");
  const inputSheet = base.getSheetByName("入稿用");
  if (!inputSheet) throw new Error("「入稿用」シートが見つかりません。");
  const masterSheet = base.getSheetByName("マスタ");
  if (!masterSheet) throw new Error("「マスタ」シートが見つかりません。");

  const registerAd = String(inputSheet.getRange("B4").getValue()).toLowerCase() === 'true';
  const onlyAd = String(inputSheet.getRange("B5").getValue()).toLowerCase() === 'true';
  const onlyMaterial = String(inputSheet.getRange("B6").getValue()).toLowerCase() === 'true';

  const modeFlags = [registerAd, onlyAd, onlyMaterial].filter(flag => flag);
  if (modeFlags.length !== 1) {
    throw new Error("登録モード（B4:広告登録＋素材追加, B5:広告登録のみ, B6:素材追加のみ）のいずれか1つのみをチェックしてください。");
  }

  let promotionId = null;
  let promotionName = null;

  if (registerAd || onlyAd) {
    Logger.log("広告登録処理開始");
    const advertiserName = inputSheet.getRange("A2").getValue();
    Logger.log("広告主名: " + advertiserName);

    // マスタから広告主ID取得
    const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 6).getValues();
    const advertiser = masterData.find(row => row[0] === advertiserName);
    Logger.log("マスタ広告主検索結果: " + JSON.stringify(advertiser));
    if (!advertiser) throw new Error("広告主がマスタに見つかりません: " + advertiserName);
    const advertiserId = advertiser[3];
    Logger.log("広告主ID: " + advertiserId);

    promotionName = entrySheet.getRange("A2").getValue();
    Logger.log("案件名: " + promotionName);
    const promotionCategoryLabel = entrySheet.getRange("B2").getValue();

    const categoryMap = {
      "エンタメ・ゲーム": "PC001",
      "Webサービス": "PC002",
      "インターネット": "PC003",
      "EC・物販": "PC004",
      "健康・美容・ファッション": "PC005",
      "グルメ・食品": "PC006",
      "お小遣い・ポイント": "PC007",
      "旅行・交通": "PC008",
      "金融・投資・保険": "PC009",
      "暮らし・不動産": "PC010",
      "仕事・学び・資格": "PC011",
      "ギフト・プレゼント": "PC012",
      "スポーツ・趣味": "PC013",
      "結婚・恋愛・出会い": "PC014",
      "その他": "PC015"
    };

    const promotionCategoryId = categoryMap[promotionCategoryLabel];
    Logger.log("広告カテゴリラベル: " + promotionCategoryLabel + ", ID: " + promotionCategoryId);
    if (!promotionCategoryId) {
      throw new Error("不正な広告カテゴリー: " + promotionCategoryLabel);
    }

    // デバイス
    const devices = [
      entrySheet.getRange("D2").getValue(),
      entrySheet.getRange("D3").getValue(),
      entrySheet.getRange("D4").getValue(),
      entrySheet.getRange("D5").getValue()
    ];
    const deviceIds = ["MM001", "MM002", "MM003", "MM004"];
    const selectedDeviceIds = deviceIds.filter((_, i) => String(devices[i]).toLowerCase() === 'true');
    Logger.log("選択デバイスID: " + JSON.stringify(selectedDeviceIds));
    if (selectedDeviceIds.length === 0) throw new Error("対応デバイスは1つ以上選択してください。");

    // 保証
    const guarantee = entrySheet.getRange("E2").getValue();
    const grossCost = entrySheet.getRange("F2").getValue();
    const netCost = entrySheet.getRange("G2").getValue();

    // 丸め
    const rounding = [
      entrySheet.getRange("I2").getValue(),
      entrySheet.getRange("I3").getValue(),
      entrySheet.getRange("I4").getValue()
    ];
    const roundingTypes = [0, 1, 2];
    const selectedRounding = roundingTypes.filter((_, i) => String(rounding[i]).toLowerCase() === 'true');
    Logger.log("選択丸め誤差: " + JSON.stringify(selectedRounding));
    if (selectedRounding.length !== 1) throw new Error("丸め誤差は1つだけ選択してください。");

    // 成果有効期間
    const validityPeriod = entrySheet.getRange("B8").getValue();

    // 成果承認
    const approval = [entrySheet.getRange("D8").getValue(), entrySheet.getRange("D9").getValue()];
    const approvalTypes = [1, 0];
    const selectedApproval = approvalTypes.filter((_, i) => String(approval[i]).toLowerCase() === 'true');
    Logger.log("選択成果承認: " + JSON.stringify(selectedApproval));
    if (selectedApproval.length !== 1) throw new Error("成果承認設定は1つだけ選択してください。");

    // URL・条件
    const checkUrl = entrySheet.getRange("E8").getValue();
    const conditions = entrySheet.getRange("F8").getValue();

    // 提携承認
    const tieUp = [entrySheet.getRange("H8").getValue(), entrySheet.getRange("H9").getValue()];
    const tieUpTypes = [1, 0];
    const selectedTieUp = tieUpTypes.filter((_, i) => String(tieUp[i]).toLowerCase() === 'true');
    Logger.log("選択提携承認: " + JSON.stringify(selectedTieUp));
    if (selectedTieUp.length !== 1) throw new Error("提携承認設定は1つだけ選択してください。");

    // 既存広告チェック（同じ案件名・広告主IDの広告があれば再利用）
    const token = 'agqnoournapf:1kvu9dyv1alckgocc848socw';
    const searchUrl = 'https://otonari-asp.com/api/v1/m/promotion/search?name=' + encodeURIComponent(promotionName) + '&advertiser=' + advertiserId;
    const searchOptions = {
      method: 'get',
      headers: { 'X-Auth-Token': token },
      muteHttpExceptions: true
    };
    try {
      const searchResponse = UrlFetchApp.fetch(searchUrl, searchOptions);
      Logger.log("広告検索APIレスポンス: " + searchResponse.getContentText());
      const searchResult = JSON.parse(searchResponse.getContentText());
      if (searchResult.records && Array.isArray(searchResult.records)) {
        // 案件名・広告主IDが完全一致するもの優先
        const exact = searchResult.records.find(r => r.name === promotionName && r.advertiser === advertiserId);
        if (exact) {
          Logger.log("既存広告が存在するため再利用: " + JSON.stringify(exact));
          return { promotionName: promotionName, promotionId: exact.id };
        }
      }
    } catch (e) {
      Logger.log("広告検索APIエラー: " + e);
      // 検索エラー時は新規作成に進む
    }

    const now = Math.floor(new Date().getTime() / 1000);
    const payload = {
      advertiser: advertiserId,
      name: promotionName,
      promotion_category: promotionCategoryId,
      promotion_device: selectedDeviceIds,
      promotion_type: [guarantee ? "action" : "click"],
      net_click_cost: guarantee ? null : netCost,
      gross_click_cost: guarantee ? null : grossCost,
      net_action_cost: guarantee ? netCost : null,
      gross_action_cost: guarantee ? grossCost : null,
      net_action_cost_type: "yen",
      gross_action_cost_type: "yen",
      action_name: promotionName,
      action_cost_round: selectedRounding[0],
      action_time: validityPeriod,
      action_time_type: 86400,
      action_apply_state: selectedApproval[0],
      display_url: checkUrl,
      display_action: conditions,
      apply_state: selectedTieUp[0],
      display_date_unix: now,
      action_double_state: 0,
      action_double_type: [],
      action_double_time: 0,
      action_double_time_type: 0,
      action_ip_auth: 0,
      action_cid_del: 1,
      opens: 1,
      ip_allow: 0,
      given_state: 0,
      tier_state: 0,
      track_session_state: 1
    };

    Object.keys(payload).forEach(key => {
      if (payload[key] === null || payload[key] === undefined || payload[key] === "") {
        delete payload[key];
      }
    });

    Logger.log("広告登録API payload: " + JSON.stringify(payload));

    const url = 'https://otonari-asp.com/api/v1/m/promotion/regist';
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'X-Auth-Token': token
      },
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      Logger.log("広告登録APIレスポンス: " + response.getContentText());
      const result = JSON.parse(response.getContentText());
      promotionId = result.record && result.record.id ? result.record.id : undefined;
      Logger.log("登録された広告ID: " + promotionId);
    } catch (e) {
      Logger.log("広告登録APIエラー: " + e);
      throw e;
    }
    Logger.log("広告登録処理終了");
  }

  if (registerAd || onlyAd) {
    Logger.log("registerAdFromSheet return: promotionName=" + promotionName + ", promotionId=" + promotionId);
    return { promotionName: promotionName, promotionId: promotionId };
  }
}
