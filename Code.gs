function registerCalendarEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const calendar = CalendarApp.getDefaultCalendar();

  // データ範囲（3行目以降、B列～F列の5列分）
  const data = sheet.getRange(3, 2, sheet.getLastRow() - 2, 5).getValues(); 

  data.forEach((row, i) => {
    const rowIndex = i + 3; // 実際の行番号
    const [date, title, startTime, endTime, description] = row;

    // 入力チェック
    if (!date || !title || !startTime || !endTime) {
      console.log(`⚠️ スキップ：${rowIndex}行目 → データ未入力`);
      return;
    }

    // 日時結合（date: 日付, startTime/endTime: 時刻オブジェクト）
    const startDateTime = new Date(date);
    startDateTime.setHours(startTime.getHours());
    startDateTime.setMinutes(startTime.getMinutes());

    const endDateTime = new Date(date);
    endDateTime.setHours(endTime.getHours());
    endDateTime.setMinutes(endTime.getMinutes());

    console.log(`✅ 登録：${rowIndex}行目 → ${title} ${startDateTime}〜${endDateTime}`);

    calendar.createEvent(title, startDateTime, endDateTime, {
      description: description || ""
    });
  });

  console.log("✅ 全イベント処理が完了しました！");
}
