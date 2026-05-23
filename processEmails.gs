function processEmails() {
  const rules = [
    {
      subject: 'PRM系統帳號未登入使用',
      forward: [], label: null, trash: true
    },
    {
      subject: '**富林環球保經總公司',
      forward: ['tp**@leishan.com.tw'],
      label: '安達', trash: false
    },
    {
      subject: 'LS_新契約受理報表',
      forward: ['all-***istant@leishan.com.tw'],
      label: '安達', trash: false
    },
    {
      subject: 'LS_新契約製發清冊報表',
      forward: ['all-***istant@leishan.com.tw'],
      label: '安達', trash: false
    },
    {
      subject: '新安東京海上產物保險股份有限公司',
      forward: ['all-***istant@leishan.com.tw'],
      label: '新安東京', trash: false
    },
    {
      subject: 'A836每日開會清單_NEW',
      forward: ['all-***istant@leishan.com.tw'],
      label: '安聯', trash: false
    },
    {
      subject: '【第一金人壽】-磊山保經照會',
      forward: ['all-***istant@leishan.com.tw'],
      label: '第一金', trash: false
    },
    {
      subject: '業務員帳號定期刪除記錄報表',
      forward: [], label: null, trash: true
    },
    {
      subject: '保誠人壽通知-磊山(BR030)發單檔',
      forward: ['all-***istant@leishan.com.tw'],
      label: '保誠', trash: false
    },
    {
      subject: '新光人壽銀行保代新契約補辦通知',
      forward: ['all-***istant@leishan.com.tw'],
      label: '新光', trash: false
    },
    {
      subject: '【安達人壽】業務員資格暨聯絡資料通知',
      forward: ['xi****@leishan.com.tw'],
      label: '安達', trash: false
    },
    {
      subject: '【安達人壽】0800電話諮詢服務通知',
      forward: ['08**@leishan.com.tw'],
      label: '安達', trash: false
    },
    {
      subject: '【安達人壽】外撥電訪名單通知',
      forward: [],
      label: '安達', trash: false
    },
    {
      subject: '新光人壽保代新契約補辦件彙總通知函[磊山保經]共',
      forward: [], label: null, trash: true
    },
    {
      subject: '全球人壽業務員自動匯入結果通知',
      forward: ['wen****lu@leishan.com.tw'],
      label: '全球', trash: false
    },
    {
      from: 'skl.eservice@skl.com.tw',
      subject: '保單簽收單回條通知',
      forward: ['all-***istant@leishan.com.tw', 'y****lin@leishan.com.tw'],
      label: '新光', trash: false
    },
    {
      subject: '安聯人壽_保單貸款明細總覽',
      forward: [],
      label: '安聯', trash: false
    },
    {
      subject: '新契約保費請款進度控制表',
      forward: [],
      label: '安聯', trash: false
    },
    {
      subject: 'A836每日照會清單_NEW',
      forward: ['all-***istant@leishan.com.tw'],
      label: '安聯', trash: false
    },
    {
      from: 'admin@fglife.com.tw',
      subject: '有不全照會通知',
      forward: [], label: null, trash: true
    },
    {
      from: '3500100@fglife.com.tw',
      subject: '銀行扣款不成功報表通知',
      forward: [], label: null, trash: true
    },
    {
      from: '3500100@fglife.com.tw',
      subject: '應收件明細表報表產生',
      forward: [], label: null, trash: true
    },
    {
      from: '3500100@fglife.com.tw',
      subject: '自動轉帳(信用卡)扣款不成功報表通知',
      forward: [], label: null, trash: true
    },
    {
      from: 'admin@fglife.com.tw',
      subject: '自行上遠雄官網註冊成功',
      forward: ['qing******e@leishan.com.tw'],
      label: '遠雄', trash: false
    },
    {
      from: 'admin@fglife.com.tw',
      subject: '寄送密碼',
      forward: ['qing******e@leishan.com.tw'],
      label: '遠雄', trash: false
    },
    {
      subject: '凱基人壽理賠受理通知',
      forward: ['all-***istant@leishan.com.tw'],
      label: '凱基', trash: false
    },
    {
      subject: '凱基人壽理賠結案通知',
      forward: ['all-***istant@leishan.com.tw'],
      label: '凱基', trash: false
    },
    {
      subject: '凱基人壽理賠照會',
      forward: ['all-***istant@leishan.com.tw'],
      label: '凱基', trash: false
    },
    {
      subject: '凱基人壽核保照會',
      forward: ['all-***istant@leishan.com.tw'],
      label: '凱基', trash: false
    },
    {
      subject: '凱基人壽保服照會',
      forward: ['all-***istant@leishan.com.tw'],
      label: '凱基', trash: false
    },
  ];

  rules.forEach(rule => {
    // 建立搜尋條件
    let query = `in:inbox subject:"${rule.subject}"`;
    if (rule.from) query += ` from:${rule.from}`;

    const threads = GmailApp.search(query);
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        // 轉寄
        rule.forward.forEach(addr => {
          GmailApp.sendEmail(addr, message.getSubject(), '', {
            htmlBody: message.getBody(),
            attachments: message.getAttachments(),
            from: Session.getActiveUser().getEmail()
          });
        });

        // 套用標籤
        if (rule.label) {
          const label = GmailApp.getUserLabelByName(rule.label);
          if (label) thread.addLabel(label);
        }

        // 刪除或歸檔
        if (rule.trash) {
          message.moveToTrash();
        } else {
          thread.moveToArchive();
        }
      });
    });
  });
}
