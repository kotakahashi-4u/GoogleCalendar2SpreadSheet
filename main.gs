function getNewAppointmentDetails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('スプレッドシートのシート名を記載'),
        calendar = CalendarApp.getCalendarById('ご自身のGoogleアカウントを記載'),
        properties = PropertiesService.getScriptProperties(),
        lastChecked = properties.getProperty('lastChecked');  // GASのスクリプトプロパティを用いて前回実施日時を取得
  
  const now = new Date(),
        startTime = lastChecked ? new Date(parseInt(lastChecked)) : new Date(now.getTime() - 5 * 60 * 1000),
        events = calendar.getEvents(startTime, new Date(now.getTime() + 24 * 60 * 60 * 1000 * 30)); // 30日後まで検索

  // 取得したイベントは一旦すべて走査対象
  events.forEach(function(event) {
    // イベントの作成日時が前回のチェック日時より後のものだけをスプレッドシートに書き込む
    if (event.getDateCreated() > startTime) {
      // スプレッドシートに行を追加
      sheet.appendRow([
        new Date(event.getStartTime()),
        new Date(event.getEndTime()),
        event.getTitle(),
        event.getGuestList().length > 0 ? event.getGuestList()[0].getEmail() : ''
      ]);
    }
  });

  // 今回の実行日時をスクリプトプロパティに保存
  properties.setProperty('lastChecked', now.getTime().toString());
}
