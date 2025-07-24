function fetchMonthlyResultsAndWrite() {
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1") 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("月次成果") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("月次成果");
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 8).setValues([
    ["成果ID", "広告名", "広告ID", "成果内容", "成果発生日時", "状態", "グロス報酬", "ネット報酬"]
  ]);

  const accessKey = 'agqnoournapf';
  const secretKey = '1kvu9dyv1alckgocc848socw';
  const token = `${accessKey}:${secretKey}`;
  const startDate = inputSheet.getRange("A2").getValue();
  const endDate = inputSheet.getRange("A3").getValue();

  const baseUrl = 'https://otonari-asp.com/api/v1/m/action_log_raw/search';
  const startUnix = Math.floor(new Date(startDate).getTime() / 1000);
  const endUnix = Math.floor(new Date(endDate).getTime() / 1000);
  const limit = 100;
  let offset = 0;
  const allResults = [];

  while (true) {
    const url = `${baseUrl}?regist_unix_start=${startUnix}&regist_unix_end=${endUnix}&offset=${offset}&limit=${limit}`;
    const options = {
      method: 'get',
      headers: {
        'X-Auth-Token': token
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log(`❌ APIエラー: ${code}`);
      break;
    }

    const data = JSON.parse(response.getContentText());
    if (!data.records || data.records.length === 0) break;

    allResults.push(...data.records);
    if (data.records.length < limit) break;

    offset += limit;
  }

  const output = allResults.map(result => {
    const id = result.id || "";
    const promotionName = result.promotion_name || "";
    const promotionId = result.promotion || "";
    const subject = result.subject || "";
    const registDate = result.regist_unix ? new Date(result.regist_unix * 1000).toLocaleString() : "";
    const state = result.state === 1 ? "承認" : result.state === 2 ? "キャンセル" : "未承認";
    const gross = result.gross_reward || 0;
    const net = result.net_reward || 0;
    return [id, promotionName, promotionId, subject, registDate, state, gross, net];
  });

  if (output.length > 0) {
    sheet.getRange(2, 1, output.length, 8).setValues(output);
    Logger.log(`✅ ${output.length} 件の成果を取得して書き込みました。`);
  } else {
    Logger.log("⚠️ 対象期間の成果が見つかりませんでした。");
  }
}
