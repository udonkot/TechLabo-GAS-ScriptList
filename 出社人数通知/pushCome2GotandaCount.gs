/**
 * �o�Аl���ʒm�X�v���b�h�V�[�g����X�N���v�g
 * �X�v���b�h�V�[�g����l���擾���ă��b�Z�[�W�𐶐��ASlack�֒ʒm
 */

// Slack�g�[�N���A�v���p�e�B����擾
const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN");

// �X�v���b�h�V�[�g���A�v���p�e�B����擾
const SHEET_URL = PropertiesService.getScriptProperties().getProperty("SPREADSHEETS_URL");

// SlackAPI
const POST_URL = PropertiesService.getScriptProperties().getProperty("SLACK_APIURL_POSTMESSAGE");

// Slack ChannelID
// bot_test
const CHANNEL_ID = PropertiesService.getScriptProperties().getProperty("SLACK_CHANNELID_BOTTEST");
// cmn-�{��7f�o�ދΘA��
// const CHANNEL_ID = PropertiesService.getScriptProperties().getProperty("SLACK_CHANNELID_COME2GOTANDA");

/**
 * main function
 */
function pushCome2GotandaCount() {

  let sendMessage = '';

  try {

    // Spreadsheet�t�@�C�����J��
    var spreadSheet = SpreadsheetApp.openByUrl(SHEET_URL);

    // �V�X�e�����t�擾
    today = new Date();

    // �V�[�g���擾
    const sheetName = getSheetName(today);
    // Sheet���J��
    var sheet = spreadSheet.getSheetByName(sheetName);

    // �����擾
    const nextDay = getNextWorkDay(today);

    // �����̗�ԍ��擾
    let targetCol = 0;
    // 1�񂲂Ƃ̂ɔ���A3����������ł���΂悢�̂�100�ŉ��u���B
    for(let i=3; i < 100; i++) {
      // ���t�Z���擾
      const dateCell = sheet.getRange(15,i);

      // �ʒm�Ώۂ̓��t������
      if(dateCell.getValue().toString() == nextDay.toString()) {
        targetCol = i;
        break;
      }

      // ��̃Z����������I��
      if(dateCell == undefined || dateCell.getValue() == '') {
        break;
      }
    }

    // ***************
    // �o�̓��b�Z�[�W����
    // ***************
    const header = '�y' + nextDay.getFullYear() + '/' + (nextDay.getMonth()+1) + '/' + nextDay.getDate() + '�̏o�Аl���z';
    const earth = createSummaryMessage(sheet, 4, targetCol);
    const venus = createSummaryMessage(sheet, 5, targetCol);
    const mars = createSummaryMessage(sheet, 6, targetCol);
    const mercury = createSummaryMessage(sheet, 7, targetCol);
    const sixA = createSummaryMessage(sheet, 8, targetCol);
    const sixB = createSummaryMessage(sheet, 9, targetCol);
    const sixZ = createSummaryMessage(sheet, 10, targetCol);

    // ���s�R�[�h�����Đ��`
    sendMessage = header + '\n' + earth + '\n' + venus + '\n' + mars + '\n' + mercury + '\n' + sixA + '\n' + sixB + '\n' + sixZ

  } catch(e) {
    Logger.log(e);
    throw new Error("SpreadSheet Error:");
  }

  try {
    // slack��post����f�[�^����
    const res = sendMessage;
    const options = getSlackConnectData(CHANNEL_ID, res);
    Logger.log(options);

    // slack�ɒʒm
    return UrlFetchApp.fetch(POST_URL, options);
    // return true
  } catch(e) {
    Logger.log(e);
    throw new Error("UrlFetchApp Error:");
  }
}

/**
 * �W�v�p���b�Z�[�W����
 */
const createSummaryMessage = (sheet, rowNumber, colNumber) => {

  // ��c����
  const roomName = sheet.getRange(rowNumber, 2).getValue();
  // �l��
  const count = sheet.getRange(rowNumber, colNumber).getValue().toString(); 
  // 
  return roomName + ':' + count + '�l'

}

/**
 * Slack���M�I�u�W�F�N�g����
 */
const getSlackConnectData = (channel, res) => {
  var payload = {
    token   : SLACK_TOKEN,
    channel : channel,
    text    : res,
  }

  var options = {
    'method' : 'post',
    // 'contentType': 'application/json',
    // 'payload': JSON.stringify(payload)
    'payload': payload
  }

  return options;
 
}

/**
 * TODO�F�V�[�g���擾
 * �Q�Ƃ���V�[�g�����擾����
 * 
 * nowDate:�V�X�e�����t�̃I�u�W�F�N�g
 */
const getSheetName = (nowDate) => {
  // �V�X�e�����t����N�����擾

  // �擾���������猻��Q����肵�ĕ�����𐶐�
  // 1Q(10��~12��) �� 10-12
  // 2Q(01��~3��) �� 01-03
  // 3Q(04��~06��) �� 04-06
  // 4Q(07��~09��) �� 07-09

  // "�V�X�e���N_��L������̃V�[�g���𐶐����ĕԋp

  // FIXME: ���I�ɃV�[�g���𐶐�����悤�C�����K�v
  const sheetName = "2023_07-09"

  return sheetName;


}

/**
 * TODO�F�y���`�F�b�N
 * �V�X�e�����t���y�������肷��(true�F�y��)
 * 
 * nowDate:�V�X�e�����t�̃I�u�W�F�N�g
 * 
 */
const isHoliday = (nowDate) => {
  // 

  return false;

}

/**
 * TODO�F���c�Ɠ��擾
 * �V�X�e�����t�̗��c�Ɠ���ԋp����
 * 
 * nowDate:�V�X�e�����t�̃I�u�W�F�N�g
 * 
 */
const getNextWorkDay = (nowDate) => {

  // FIXME:�������y���̏ꍇ�͗��T���j�̓��t��ԋp����悤�C�����K�v
  const nextWorkDay = nowDate;
  nextWorkDay.setDate(nextWorkDay.getDate()+1);
  nextWorkDay.setHours(0,0,0)

  return nextWorkDay;


}

