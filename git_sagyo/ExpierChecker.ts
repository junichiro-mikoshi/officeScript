function main(workbook: ExcelScript.Workbook) {
    let ws = workbook.getActiveWorksheet();
    let lastRow: number = ws.getUsedRange().getLastRow().getRowIndex();
    let result: string = "";
    let resultAry: string[] = [];
    let today = new Date();
    
    for (let i = 2; i <= lastRow; i++) {

        // 「契約期限」セルから日付を取得し、シリアル値から日付型のデータに変換する。
        let dateValue: number = ws.getCell(i, 7).getValue();

        // 「契約期限」が空白の場合に処理をSkipする。
        if (!dateValue) {
            continue;
        }

        let deadlineDate = (dateValue);

        // 日付の差分を計算する。
        let diffDays = Math.round((deadlineDate - today) / (1000 * 60 * 60 * 24));

        // 契約期限まで45日以内である「ソフトウェア名」を 出力する。
        if (diffDays <= 0) {
            let name: string = ws.getCell(i, 2).getValue();
            resultAry.push(name + "は期限切れです。");

        } else if (diffDays <= 45) {
            let name: string = ws.getCell(i, 2).getValue();
            resultAry.push(name + "はあと" + diffDays + "日で期限が切れます。");
        }

    }
    if (resultAry.length == 0) {
        result = "45日以内で更新期限が迫っているものはありません。"
    } else {
        result = resultAry.join('\r\n');
    }
    return result;
}
