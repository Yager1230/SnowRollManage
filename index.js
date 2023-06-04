// 假设你有一个JSON数组
var data = [
  { name: "张三", age: 20, gender: "男" },
  { name: "李四", age: 21, gender: "女" },
  { name: "王五", age: 22, gender: "男", department: 'dtch' }
];

// 将JSON数组转换为工作表对象
var ws = XLSX.utils.json_to_sheet(data);

// 创建一个工作簿对象
var wb = XLSX.utils.book_new();

// 将工作表添加到工作簿中
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

// 导出并下载xlsx文件
const exportXlsxButton = document.getElementById('exportXlsxButton');
exportXlsxButton.addEventListener('click', () => {
  XLSX.writeFile(wb, "data.xlsx");
});


var reader = new FileReader();

// 定义一个回调函数，当文件读取完成后执行
reader.onload = function (e) {

  console.log(123);
  // 获取文件的二进制数据
  var data = e.target.result;
  // 解析数据为工作簿对象
  var workbook = XLSX.read(data, { type: "binary" });
  // 获取第一个工作表的名称
  var sheetName = workbook.SheetNames[0];
  // 获取第一个工作表的数据
  var sheetData = workbook.Sheets[sheetName];
  // 将数据转换为 JSON 格式
  var jsonData = XLSX.utils.sheet_to_json(sheetData);
  // 打印 JSON 数据
  console.log(jsonData);
};

const importXlsxButton = document.getElementById('importXlsxButton');
importXlsxButton.addEventListener('change', (e) => {

  // 定义一个函数，用于导入并读取 xlsx 文件
  function importAndReadXlsx(file) {
    // 检查文件是否是 xlsx 格式
    if (file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      // 读取文件为二进制字符串
      reader.readAsBinaryString(file);
    } else {
      // 提示错误信息
      alert("请导入 xlsx 格式的文件");
    }
  }


  // importAndReadXlsx();
});

const channelBox = document.getElementById('channelBox');
channelBox.addEventListener('click', () => {
  getChannelList();
});

let islandId = '';
const islandIdInpt = document.getElementById('islandIdInpt');
islandIdInpt.addEventListener('input', () => {
  islandId = islandIdInpt.value;
});

const signContentInpt = document.getElementById('signContentInpt');
signContentInpt.addEventListener('input', () => {
  myData.signMessage = signContentInpt.value;
});

// 指定需要展示的频道数据
let channelFilterName = '';
const channelNameInpt = document.getElementById('channelNameInpt');
channelNameInpt.addEventListener('input', () => {
  channelFilterName = channelNameInpt.value;
});

const signMessageBox = document.getElementById('signMessageBox');
signMessageBox.addEventListener('click', () => {
  senSignMessage();
});

const signedMemberBox = document.getElementById('signedMemberBox');
signedMemberBox.addEventListener('click', () => {
  setSignMembers();
});

const historyCheckbox = document.getElementById('historyCheckbox');
historyCheckbox.addEventListener('click', () => {
  if (historyCheckbox.checked) {
    showHistoryTable();
  } else {
    hideHistoryTable();
  }
});

const signedMessageIdInpt = document.getElementById('signedMessageIdInpt');
signedMessageIdInpt.addEventListener('input', () => {
  myData.signMessageId = signedMessageIdInpt.value;
});

const coinSelector = document.getElementById('coinSelector');
coinSelector.addEventListener('change', () => {
  myData.coninOperType = Number(coinSelector.value);

});

const setCoinTest = document.getElementById('setCoinTest');
setCoinTest.addEventListener('click', () => {
  // TODO 待支持自定义
  setSingleMemberCoin({
    dodoSourceId: myData.signedMembers[0].dodoSourceId,
    operateType: Number(coinSelector.value),
    integral: 100,
  });
});


// exportXlsxButton.addEventListener('click', () => {
//   var data = [
//     ['姓名', '性别', '年龄'],
//     ['张三', '男', '18'],
//     ['李四', '女', '22'],
//     ['王五', '男', '28']
//   ];
//   var ws = XLSX.utils.aoa_to_sheet(data);
//   var wb = XLSX.utils.book_new();
//   XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
//   XLSX.writeFile(wb, "SheetJS.xlsx");
// });


const signedTableBody = document.getElementById('signedTableBody');
const historyTable = document.getElementById('historyTable');
const messageIdContainer = document.getElementById('messageIdContainer');
const channelIdBox = document.getElementById('channelIdBox');
const tableBody = document.getElementById('tableBody');

/* 定义所有变量在这里定义 **/
const myData = {
  channelId: '',
  signMessage: 'WOW！我根本没想好要说什么。',
  signMessageId: '',
  signedMembers: [],
  channelData: {},
};
const SESSION_KEY = 'signedList';

/** 获取频道列表 */
function getChannelList() {
  sendHttpRequest({
    url: '/api/v2/channel/list',
    requestBody: {
      islandSourceId: islandId.trim(),
    },
    callback: (res) => {
      myData.channelData = res.data.find(item => item.channelName === channelFilterName.trim());
      if (!myData.channelData) {
        channelIdBox.innerHTML = '未找到相关频道，请填写正确的频道名称。';
        return;
      }
      myData.channelId = myData.channelData.channelId;
      const channelName = myData.channelData.channelName;
      channelIdBox.innerHTML = `
        <li>频道Id: ${myData.channelId}</li>
        <li>频道名称：${channelName}</li>
      `;
    },
    errorMessage: '获取频道列表失败',
  });
}

/** 发送消息 */
function senSignMessage() {
  if (!myData.signMessage) {
    window.alert('请输入要发送的信息。');
  }
  sendHttpRequest({
    url: '/api/v2/channel/message/send',
    errorMessage: '发送消息失败',
    requestBody: {
      channelId: myData.channelId,
      messageType: 1,
      messageBody: { content: myData.signMessage }
    },
    callback: (res) => {
      myData.signMessageId = res.data.messageId;
      messageIdContainer.innerHTML = `消息内容: ${myData.signMessage}   (消息id: ${res.data.messageId})`;
      const signedMessageList = JSON.parse(sessionStorage.getItem(SESSION_KEY)) || [];
      sessionStorage.setItem(SESSION_KEY, JSON.stringify([{ id: myData.signMessageId, content: myData.signMessage }, ...signedMessageList]));
    }
  });
}

/* 获取此时签到人数的快照 固定表情id 128516 **/
function setSignMembers() {
  sendHttpRequest({
    url: '/api/v2/channel/message/reaction/member/list',
    errorMessage: '获取消息反应成员列表失败',
    requestBody: {
      messageId: myData.signMessageId,
      emoji: {
        id: '128516',
        type: 1,
      },
      pageSize: 100,
      maxId: 0,
    },
    callback: (res) => {
      myData.signedMembers = [...res.data.list];
      initSignedTable();
    }
  });
}

/* 给单个成员发放雪币 **/
function setSingleMemberCoin({ dodoSourceId, operateType, integral }) {
  sendHttpRequest({
    url: '/api/v2/integral/edit',
    errorMessage: '分配成员积分失败',
    requestBody: {
      islandSourceId: islandId,
      dodoSourceId,
      operateType,
      integral,
    },
    callback: (res) => {
    }
  });
}

function showHistoryTable() {
  const sessionData = JSON.parse(sessionStorage.getItem(SESSION_KEY)) || [];
  const bodyContent = sessionData.map(item => `
    <tr>
      <td>${item.content}</td>
      <td>${item.id}</td>
    </tr>
  `).join('');
  historyTable.innerHTML = `
    <thead>
      <th>消息内容</th>
      <th>消息id</th>
    </thead>
    <tbody>
      ${bodyContent}
    </tbody>
  `;
}

function hideHistoryTable() {
  historyTable.innerHTML = null;
}


/** 接口调用 */
function sendHttpRequest({ url, callback = () => { }, requestBody, errorMessage }) {
  const DODO_URL = 'https://botopen.imdodo.com';
  const xhr = new XMLHttpRequest();
  xhr.open('POST', `${DODO_URL + url}`, true);
  xhr.setRequestHeader('Content-Type', 'application/json');
  xhr.setRequestHeader('Authorization', 'Bot 52804622.NTI4MDQ2MjI.Wu-_vXB2.EGhVavQO8NX7JGkg2F1M_sWKjcomwCq_9fUfPwOwJao');
  xhr.withCredentials = true;
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4 && xhr.status === 200) {
      const res = JSON.parse(xhr.responseText);
      if (res?.message !== 'success') {
        window.alert(`${errorMessage}: ${res?.message || '未知错误'}`);
        return;
      }
      callback(res);
    }
  };
  if (requestBody) {
    xhr.send(JSON.stringify(requestBody));
  } else {
    xhr.send();
  }
}

/** 渲染要展示的表格数据 */
function initSignedTable() {
  if (!myData.signedMembers.length) {
    window.alert('还没有人签到。');
  }
  signedTableBody.innerHTML = myData.signedMembers.map(item => (
    `<tr>
      <td>${item.nickName}</td>
      <td>${item.dodoSourceId}</td>
      <td><button>发放积分</button></td>
    </tr>`
  )).join('');
}

