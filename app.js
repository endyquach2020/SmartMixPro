// ==========================================================================
// SMARTMIX PRO - ENGINE XỬ LÝ WORD OPENXML & EXCEL MA TRẬN ĐÁP ÁN (BẢN FIX DÀN CỘT)
// ==========================================================================

let uploadedFile = null;
let generatedBlobs = [];
let answerExcelBlob = null;

// Hàm chuyển tab
function switchTab(tabName) {
  const tabMixBtn = document.getElementById('tabMixBtn');
  const tabGuideBtn = document.getElementById('tabGuideBtn');
  const tabMixContent = document.getElementById('tabMixContent');
  const tabGuideContent = document.getElementById('tabGuideContent');

  if (!tabMixBtn || !tabGuideBtn || !tabMixContent || !tabGuideContent) return;

  if (tabName === 'mix') {
    tabMixBtn.className = "tab-btn active flex items-center space-x-2 py-2 px-4 rounded-lg font-semibold text-xs sm:text-sm bg-white text-blue-700 shadow-sm border border-slate-200/60 transition";
    tabGuideBtn.className = "tab-btn inactive flex items-center space-x-2 py-2 px-4 rounded-lg font-medium text-xs sm:text-sm text-slate-600 hover:text-slate-900 hover:bg-white/60 transition";
    tabMixContent.style.display = "block";
    tabGuideContent.style.display = "none";
  } else {
    tabMixBtn.className = "tab-btn inactive flex items-center space-x-2 py-2 px-4 rounded-lg font-medium text-xs sm:text-sm text-slate-600 hover:text-slate-900 hover:bg-white/60 transition";
    tabGuideBtn.className = "tab-btn active flex items-center space-x-2 py-2 px-4 rounded-lg font-semibold text-xs sm:text-sm bg-white text-blue-700 shadow-sm border border-slate-200/60 transition";
    tabMixContent.style.display = "none";
    tabGuideContent.style.display = "block";
  }
}

// Khởi tạo các sự kiện khi DOM tải xong
document.addEventListener('DOMContentLoaded', () => {
  const tabMixBtn = document.getElementById('tabMixBtn');
  const tabGuideBtn = document.getElementById('tabGuideBtn');
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const fileNameLabel = document.getElementById('fileNameLabel');
  const btnProcess = document.getElementById('btnProcess');
  const btnDownloadAll = document.getElementById('btnDownloadAll');

  if (tabMixBtn) tabMixBtn.addEventListener('click', () => switchTab('mix'));
  if (tabGuideBtn) tabGuideBtn.addEventListener('click', () => switchTab('guide'));

  // Lưu và tải thông tin chung từ localStorage (giúp ghi nhớ cố định các thông tin đã nhập)
  const infoFieldIds = ['infoDepartment', 'infoSchool', 'infoExamName', 'infoSchoolYear', 'infoSubject', 'infoDuration'];
  infoFieldIds.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      const saved = localStorage.getItem(`smartmix_${id}`);
      if (saved) {
        el.value = saved;
      }
      el.addEventListener('input', () => {
        localStorage.setItem(`smartmix_${id}`, el.value);
      });
    }
  });

  if (dropZone && fileInput) {
    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropZone.classList.add('drag-over', 'border-blue-500', 'bg-blue-50/50');
    });

    dropZone.addEventListener('dragleave', () => {
      dropZone.classList.remove('drag-over', 'border-blue-500', 'bg-blue-50/50');
    });

    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('drag-over', 'border-blue-500', 'bg-blue-50/50');
      if (e.dataTransfer.files && e.dataTransfer.files.length) {
        processSelectedFile(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length) {
        processSelectedFile(e.target.files[0]);
      }
    });
  }

  function processSelectedFile(file) {
    if (!file.name.toLowerCase().endsWith('.docx')) {
      alert('Vui lòng chọn file Word định dạng .docx');
      return;
    }
    uploadedFile = file;
    const fileSizeKb = (file.size / 1024).toFixed(1);
    if (fileNameLabel) {
      fileNameLabel.innerHTML = `
        <div class="inline-flex items-center space-x-2 px-3 py-1.5 bg-blue-50/80 border border-blue-200 rounded-lg text-blue-900 shadow-2xs">
          <svg class="w-4 h-4 text-blue-600 shrink-0" fill="currentColor" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6zm4 18H6V4h7v5h5v11z"/></svg>
          <span class="font-bold text-xs truncate max-w-xs sm:max-w-md">${file.name}</span>
          <span class="text-[12px] text-blue-600/80 font-mono">(${fileSizeKb} KB)</span>
          <span class="text-[13px] text-emerald-600 font-bold ml-1 flex items-center">&check; Đã sẵn sàng</span>
        </div>
      `;
    }
    if (btnProcess) {
      btnProcess.disabled = false;
      btnProcess.className = "w-full h-10 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold rounded-lg text-xs sm:text-sm shadow-md shadow-blue-500/20 hover:shadow-lg transition-all duration-200 cursor-pointer flex items-center justify-center space-x-2 active:scale-98";
    }
  }

  if (btnProcess) {
    btnProcess.addEventListener('click', startProcessing);
  }

  if (btnDownloadAll) {
    btnDownloadAll.addEventListener('click', downloadAllZip);
  }
});

// Xáo trộn mảng Fisher-Yates
function shuffleArray(arr) {
  const copy = [...arr];
  for (let i = copy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

// Kiểm tra thẻ gạch chân (<w:u>) để xác định đáp án đúng
function isNodeUnderlined(node) {
  if (!node) return false;
  if (node.getElementsByTagName) {
    const uTags = node.getElementsByTagName('w:u');
    for (let u of uTags) {
      const val = u.getAttribute('w:val');
      if (val !== 'none') return true;
    }
  }
  return false;
}

// Xoá bỏ gạch chân trong file đề thi xuất ra
function removeUnderline(node) {
  if (!node || !node.getElementsByTagName) return;
  const uTags = Array.from(node.getElementsByTagName('w:u'));
  uTags.forEach(u => {
    if (u.parentNode) u.parentNode.removeChild(u);
  });
}

function removeSectionBreaks(node) {
  if (!node || !node.getElementsByTagName) return;
  const sects = Array.from(node.getElementsByTagName('w:sectPr'));
  sects.forEach(s => {
    if (s.parentNode) s.parentNode.removeChild(s);
  });
}

function cleanChoiceRuns(runs) {
  const cleaned = [];
  for (let r of runs) {
    if (r.nodeName === 'w:r') {
      const tabs = Array.from(r.getElementsByTagName('w:tab'));
      tabs.forEach(tab => {
        if (tab.parentNode) tab.parentNode.removeChild(tab);
      });
      const hasContent = (r.textContent && r.textContent.trim().length > 0) ||
        r.getElementsByTagName('w:drawing').length > 0 ||
        r.getElementsByTagName('w:object').length > 0 ||
        r.getElementsByTagName('m:oMath').length > 0;
      if (hasContent) {
        cleaned.push(r);
      }
    } else if (r.nodeName !== 'w:tab') {
      cleaned.push(r);
    }
  }
  return cleaned;
}

function isComplexExpression(runs) {
  let combinedText = '';
  for (let r of runs) {
    combinedText += (r.textContent || '');
    if (r.getElementsByTagName && r.getElementsByTagName('w:fldSimple').length > 0) return true;
    if (r.getElementsByTagName && r.getElementsByTagName('w:br').length > 0) return true;
  }
  if (/(\(\vert{}\)|\[\vert{}\]|\{|\}).*?→/.test(combinedText)) return true;
  if (/\{/.test(combinedText) && /\n/.test(combinedText)) return true;
  return false;
}

function calculateVisualWidth(runs) {
  let width = 0;
  for (let r of runs) {
    if (r.nodeName === 'm:oMath' || (r.getElementsByTagName && r.getElementsByTagName('m:oMath').length > 0)) {
      const mathText = r.textContent.trim();
      if (mathText.length <= 4) width += 4;
      else width += Math.max(mathText.length, 6);
    } else if (r.nodeName === 'w:r') {
      const tTags = r.getElementsByTagName('w:t');
      for (let t of tTags) {
        width += t.textContent.trim().length;
      }

      // Nhận diện kích thước công thức MathType OLE Object
      const shapes = r.getElementsByTagName('v:shape');
      for (let s of shapes) {
        const style = s.getAttribute('style') || '';
        const ptMatch = style.match(/width:\s*(\d+(?:\.\d+)?)pt/i);
        if (ptMatch) {
          const ptVal = parseFloat(ptMatch[1]);
          width += Math.ceil(ptVal / 6);
        } else {
          width += 8;
        }
      }

      // Nhận diện kích thước hình vẽ hoặc công thức w:drawing
      const drawings = r.getElementsByTagName('w:drawing');
      for (let d of drawings) {
        const extent = d.getElementsByTagName('wp:extent')[0];
        if (extent && extent.getAttribute('cx')) {
          const cx = parseInt(extent.getAttribute('cx'));
          width += Math.ceil(cx / 76200);
        } else {
          width += 10;
        }
      }
    }
  }
  return width;
}

// BỘ TRÍCH XUẤT CÂU TRẮC NGHIỆM PHẦN I (TÁCH THÂN CÂU, BẢNG, VÀ 4 PHƯƠNG ÁN A, B, C, D)
function parseMultipleChoiceQuestion(qNodes) {
  const choiceLabels = ['A', 'B', 'C', 'D'];
  let stemNodes = [];
  let collectedChoices = {};
  let currentLabel = null;

  for (let node of qNodes) {
    // Nếu là Bảng dữ liệu <w:tbl>, giữ nguyên vào phần thân câu hỏi
    if (node.nodeName === 'w:tbl') {
      if (Object.keys(collectedChoices).length === 0) {
        stemNodes.push(node.cloneNode(true));
      }
      continue;
    }

    if (node.nodeName !== 'w:p') continue;

    const pText = node.textContent || '';
    const hasLabelInPara = /(?:^|\s|\t)*([A-D])[\.:\)]/.test(pText);
    if (hasLabelInPara) {
      currentLabel = null;
    }

    const children = Array.from(node.childNodes).filter(n => n.nodeName !== 'w:pPr');
    let hasChoiceInThisPara = false;

    for (let child of children) {
      const txt = child.textContent || '';
      // Nhận diện nhãn phương án dạng: "A.", "B.", "C.", "D." hoặc "A:", "A)"
      const match = txt.match(/(?:^|\s|\t)*([A-D])[\.:\)]/);

      if (match && choiceLabels.includes(match[1])) {
        hasChoiceInThisPara = true;
        currentLabel = match[1];
        if (!collectedChoices[currentLabel]) {
          collectedChoices[currentLabel] = {
            label: currentLabel,
            runs: [],
            isCorrect: false
          };
        }
        if (isNodeUnderlined(child)) {
          collectedChoices[currentLabel].isCorrect = true;
        }
        collectedChoices[currentLabel].runs.push(child.cloneNode(true));
      } else if (currentLabel) {
        if (isNodeUnderlined(child)) {
          collectedChoices[currentLabel].isCorrect = true;
        }
        collectedChoices[currentLabel].runs.push(child.cloneNode(true));
      }
    }

    // Nếu đoạn văn bản này chưa có phương án nào và chưa từng gặp phương án nào, nó là phần dẫn câu hỏi
    if (!hasChoiceInThisPara && Object.keys(collectedChoices).length === 0) {
      stemNodes.push(node.cloneNode(true));
    }
  }

  // Đảm bảo trích xuất đủ cả 4 phương án A, B, C, D
  const foundKeys = Object.keys(collectedChoices);
  if (foundKeys.length === 4) {
    return {
      stemNodes: stemNodes.length > 0 ? stemNodes : [qNodes[0].cloneNode(true)],
      choices: [collectedChoices['A'], collectedChoices['B'], collectedChoices['C'], collectedChoices['D']]
    };
  }

  return null;
}

// Tách các ý a), b), c), d) trong câu Đúng/Sai (Phần II - <#g2>)
function extractSubItemsTrueFalse(qNodes) {
  const subLabels = ['a', 'b', 'c', 'd'];
  let items = [];

  for (let pNode of qNodes) {
    if (pNode.nodeName !== 'w:p') continue;
    const text = pNode.textContent.trim();
    const match = text.match(/^([a-d])\)/i);
    if (match && subLabels.includes(match[1].toLowerCase())) {
      items.push({
        label: match[1].toLowerCase(),
        isCorrect: isNodeUnderlined(pNode),
        node: pNode
      });
    }
  }
  return items.length === 4 ? items : null;
}

// Trích xuất đáp án câu trả lời ngắn (Phần III - <#g3>)
// Quy tắc: Trong nhóm <#g3>, sau ký tự A. đọc hết các ký tự bên phải vì đó là đáp án
function extractShortAnswer(qNodes, isG3Section = false) {
  if (!qNodes || !qNodes.length) return null;

  // Nếu không phải trong nhóm <#g3> đã khai báo, kiểm tra an toàn tránh nhầm trắc nghiệm có B, C, D
  if (!isG3Section) {
    for (let pNode of qNodes) {
      if (pNode.nodeName !== 'w:p') continue;
      const text = (pNode.textContent || '').trim();
      if (/^\s*[B-D][\.:\)]/.test(text)) {
        return null;
      }
    }
  }

  // Ưu tiên 1: Quét từ dưới lên các đoạn bắt đầu bằng A. (ví dụ: "A. 19.", "A.19.", "Đáp án: A. 19"...)
  for (let i = qNodes.length - 1; i >= 0; i--) {
    const pNode = qNodes[i];
    if (pNode.nodeName !== 'w:p') continue;
    const text = (pNode.textContent || '').trim();

    const startMatch = text.match(/^\s*(?:đáp\s*án\s*:?|đ\/a\s*:?|đs\s*:?)?\s*A\.\s*(.*)$/i);
    if (startMatch) {
      // Đọc hết các ký tự bên phải của A. (chỉ bỏ khoảng trắng thừa đầu/cuối chuỗi, giữ nguyên toàn bộ ký tự)
      const cleanAnswer = startMatch[1].trim();
      return {
        answerText: cleanAnswer,
        nodeIndex: i,
        isInline: false
      };
    }
  }

  // Ưu tiên 2: Quét từ dưới lên các đoạn có ký tự A. ở bất kỳ vị trí nào
  for (let i = qNodes.length - 1; i >= 0; i--) {
    const pNode = qNodes[i];
    if (pNode.nodeName !== 'w:p') continue;
    const text = (pNode.textContent || '').trim();

    const inlineMatch = text.match(/(?:^|\s|\t)(?:đáp\s*án\s*:?|đ\/a\s*:?|đs\s*:?)?\s*A\.\s*(.*)$/i);
    if (inlineMatch) {
      const cleanAnswer = inlineMatch[1].trim();
      const isInline = (text.length > inlineMatch[0].trim().length + 5);
      return {
        answerText: cleanAnswer,
        nodeIndex: i,
        isInline: isInline
      };
    }
  }

  return null;
}

// Xóa bỏ phần A. [đáp án] khi nằm chung dòng với thân câu hỏi trong đề xuất ra
function removeShortAnswerFromNode(node) {
  if (!node || !node.getElementsByTagName) return;
  const textNodes = Array.from(node.getElementsByTagName("w:t"));
  let fullText = textNodes.map(t => t.textContent).join("");
  const match = fullText.match(/(?:^|\s|\t)(?:đáp\s*án\s*:?|đ\/a\s*:?|đs\s*:?)?\s*A\..*$/i);
  if (!match) return;

  const targetIndex = match.index;
  let currentPos = 0;
  for (let t of textNodes) {
    const len = t.textContent.length;
    const startPos = currentPos;
    const endPos = currentPos + len;
    currentPos = endPos;

    if (startPos >= targetIndex) {
      t.textContent = "";
    } else if (targetIndex < endPos) {
      t.textContent = t.textContent.substring(0, targetIndex - startPos).trimEnd();
    }
  }
}


// Tách riêng nhãn A., B., C., D. để chỉ in đậm nhãn
function splitAndFormatChoiceLabel(runs, newLabel) {
  const labelRegex = /(?:^|\s|\t)*([A-D])[\.:\)]\s*/;
  let found = false;

  for (let i = 0; i < runs.length; i++) {
    const r = runs[i];
    if (r.nodeName !== 'w:r') continue;

    const tTags = r.getElementsByTagName('w:t');
    for (let j = 0; j < tTags.length; j++) {
      const t = tTags[j];
      const match = t.textContent.match(labelRegex);
      if (match) {
        const afterText = t.textContent.substring(match.index + match[0].length);

        t.textContent = `${newLabel}. `;
        t.setAttribute("xml:space", "preserve");

        let rPr = r.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
          rPr = r.ownerDocument.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:rPr");
          r.insertBefore(rPr, r.firstChild);
        }
        if (!rPr.getElementsByTagName("w:b").length) {
          const b = r.ownerDocument.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:b");
          rPr.appendChild(b);
        }

        if (afterText.length > 0) {
          const afterRun = r.cloneNode(true);
          const afterRPr = afterRun.getElementsByTagName("w:rPr")[0];
          if (afterRPr) {
            const b = afterRPr.getElementsByTagName("w:b")[0];
            if (b) afterRPr.removeChild(b);
          }
          const afterT = afterRun.getElementsByTagName("w:t")[0];
          if (afterT) afterT.textContent = afterText;

          runs.splice(i + 1, 0, afterRun);
        }

        found = true;
        break;
      }
    }
    if (found) break;
  }

  // Nếu không tìm thấy nhãn trong run, chèn thêm nhãn mới in đậm vào đầu
  if (!found && runs.length > 0) {
    const doc = runs[0].ownerDocument;
    const labelRun = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:r");
    const rPr = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:rPr");
    const b = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:b");
    rPr.appendChild(b);
    labelRun.appendChild(rPr);
    const t = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = `${newLabel}. `;
    labelRun.appendChild(t);
    runs.unshift(labelRun);
  }
}

// DÀN CỘT CHUẨN XÁC: 4 CỘT, 2 CỘT HOẶC 1 CỘT THEO ĐỘ DÀI
function generateChoicesXml(choices, shouldShuffleChoices) {
  let finalChoices = choices.map(c => ({
    label: c.label,
    isCorrect: c.isCorrect,
    runs: c.runs.map(r => r.cloneNode(true))
  }));

  if (shouldShuffleChoices) {
    finalChoices = shuffleArray(finalChoices);
  }

  const labels = ['A', 'B', 'C', 'D'];
  let hasComplex = false;
  let maxVisualWidth = 0;
  let correctChoiceLetter = '';

  finalChoices.forEach((c, idx) => {
    c.newLabel = labels[idx];
    if (c.isCorrect) {
      correctChoiceLetter = c.newLabel;
    }

    c.runs.forEach(r => removeUnderline(r));
    c.runs = cleanChoiceRuns(c.runs);
    splitAndFormatChoiceLabel(c.runs, c.newLabel);

    if (isComplexExpression(c.runs)) hasComplex = true;
    const w = calculateVisualWidth(c.runs);
    if (w > maxVisualWidth) maxVisualWidth = w;
  });

  const serializer = new XMLSerializer();
  const choiceXmls = finalChoices.map(c =>
    c.runs.map(r => serializer.serializeToString(r)).join('')
  );

  const pPrSpacing = `<w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/>`;
  let xmlString = '';

  // 1. DÀN 4 CỘT TRÊN 1 DÒNG DUY NHẤT:
  // Áp dụng khi đáp án ngắn gọn (như số, tọa độ ngắn, công thức ngắn <= 18 ký tự)
  if (!hasComplex && maxVisualWidth <= 18) {
    xmlString = `
    <w:p>
      <w:pPr>
        ${pPrSpacing}
        <w:tabs>
          <w:tab w:val="left" w:pos="2551"/>
          <w:tab w:val="left" w:pos="5102"/>
          <w:tab w:val="left" w:pos="7654"/>
        </w:tabs>
      </w:pPr>
      ${choiceXmls[0]}
      <w:r><w:tab/></w:r>
      ${choiceXmls[1]}
      <w:r><w:tab/></w:r>
      ${choiceXmls[2]}
      <w:r><w:tab/></w:r>
      ${choiceXmls[3]}
    </w:p>`;
  }
  // 2. DÀN 2 CỘT TRÊN 2 DÒNG:
  // Áp dụng cho tọa độ vector, biểu thức vừa phải, hoặc câu ngắn <= 42 ký tự
  else if (!hasComplex && maxVisualWidth <= 42) {
    xmlString = `
    <w:p>
      <w:pPr>
        ${pPrSpacing}
        <w:tabs><w:tab w:val="left" w:pos="5102"/></w:tabs>
      </w:pPr>
      ${choiceXmls[0]}
      <w:r><w:tab/></w:r>
      ${choiceXmls[1]}
    </w:p>
    <w:p>
      <w:pPr>
        ${pPrSpacing}
        <w:tabs><w:tab w:val="left" w:pos="5102"/></w:tabs>
      </w:pPr>
      ${choiceXmls[2]}
      <w:r><w:tab/></w:r>
      ${choiceXmls[3]}
    </w:p>`;
  }
  // 3. DÀN 1 CỘT (khi là câu lý thuyết dài > 42 ký tự hoặc biểu thức phức tạp)
  else {
    xmlString = choiceXmls.map(xml => `
      <w:p>
        <w:pPr>${pPrSpacing}</w:pPr>
        ${xml}
      </w:p>
    `).join('');
  }

  return { xmlString, correctChoiceLetter };
}

function renumberQuestion(pNode, newIndex) {
  const qRegex = /^(Câu|Question)\s+\d+[\.:]?\s*/i;
  const textNodes = pNode.getElementsByTagName("w:t");

  for (let i = 0; i < textNodes.length; i++) {
    const t = textNodes[i];
    const match = t.textContent.match(qRegex);
    if (match) {
      const parentRun = t.closest("w\\:r") || t.parentNode;
      const afterText = t.textContent.substring(match[0].length);

      t.textContent = `${match[1]} ${newIndex}. `;
      t.setAttribute("xml:space", "preserve");

      if (parentRun && parentRun.nodeName === 'w:r') {
        let rPr = parentRun.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
          rPr = pNode.ownerDocument.createElement("w:rPr");
          parentRun.insertBefore(rPr, parentRun.firstChild);
        }
        if (!rPr.getElementsByTagName("w:b").length) {
          rPr.appendChild(pNode.ownerDocument.createElement("w:b"));
        }

        if (afterText.length > 0) {
          const afterRun = parentRun.cloneNode(true);
          const afterRPr = afterRun.getElementsByTagName("w:rPr")[0];
          if (afterRPr) {
            const b = afterRPr.getElementsByTagName("w:b")[0];
            if (b) afterRPr.removeChild(b);
          }
          const afterT = afterRun.getElementsByTagName("w:t")[0];
          if (afterT) afterT.textContent = afterText;

          parentRun.parentNode.insertBefore(afterRun, parentRun.nextSibling);
        }
      }
      break;
    }
  }
}

function boldSubQuestions(node) {
  const subRegex = /^([a-d]\))\s*/i;
  const textNodes = node.getElementsByTagName("w:t");

  for (let i = 0; i < textNodes.length; i++) {
    const t = textNodes[i];
    const match = t.textContent.match(subRegex);
    if (match) {
      const parentRun = t.closest("w\\:r") || t.parentNode;
      const afterText = t.textContent.substring(match[0].length);

      t.textContent = `${match[1]} `;
      t.setAttribute("xml:space", "preserve");

      if (parentRun && parentRun.nodeName === 'w:r') {
        let rPr = parentRun.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
          rPr = node.ownerDocument.createElement("w:rPr");
          parentRun.insertBefore(rPr, parentRun.firstChild);
        }
        if (!rPr.getElementsByTagName("w:b").length) {
          rPr.appendChild(node.ownerDocument.createElement("w:b"));
        }

        if (afterText.length > 0) {
          const afterRun = parentRun.cloneNode(true);
          const afterRPr = afterRun.getElementsByTagName("w:rPr")[0];
          if (afterRPr) {
            const b = afterRPr.getElementsByTagName("w:b")[0];
            if (b) afterRPr.removeChild(b);
          }
          const afterT = afterRun.getElementsByTagName("w:t")[0];
          if (afterT) afterT.textContent = afterText;

          parentRun.parentNode.insertBefore(afterRun, parentRun.nextSibling);
        }
      }
      break;
    }
  }
}

function cleanGroupTag(node) {
  const textNodes = node.getElementsByTagName("w:t");
  for (let i = 0; i < textNodes.length; i++) {
    textNodes[i].textContent = textNodes[i].textContent.replace(/<#g\s*[0-4]\s*>\s*/gi, '');
  }
}

// Nhận diện nhóm phần thi thông minh (kết hợp tiêu đề chuẩn BGD và thẻ <#g>)
function detectSectionType(node) {
  if (!node) return null;
  const text = (node.textContent || '').trim();
  if (!text) return null;

  // 1. Nhận diện theo từ khóa Tiêu đề (Chuẩn Bộ GD&ĐT: Phần I, Phần II, Phần III, Phần IV)
  // Ưu tiên cực cao vì tên đề mục phản ánh chính xác nội dung câu hỏi
  if (/(?:PHẦN|Phần)\s*(?:I\b|1\b)|nhiều phương án|nhiều lựa chọn/i.test(text)) {
    return 'g1';
  }
  if (/(?:PHẦN|Phần)\s*(?:II\b|2\b)|đúng\s*sai/i.test(text)) {
    return 'g2';
  }
  if (/(?:PHẦN|Phần)\s*(?:III\b|3\b)|trả lời ngắn/i.test(text)) {
    return 'g3';
  }
  if (/(?:PHẦN|Phần)\s*(?:IV\b|4\b)|tự luận/i.test(text)) {
    return 'g4';
  }

  // 2. Nhận diện theo thẻ <#g1> đến <#g4> (cho phép linh hoạt khoảng trắng như <#g 1>, <#g1 >)
  const tagMatch = text.match(/<#g\s*([1-4])\s*>/i);
  if (tagMatch) {
    return 'g' + tagMatch[1];
  }

  return null;
}

function applySpacing(pNode) {
  if (pNode.nodeName !== 'w:p') return;
  let pPr = pNode.getElementsByTagName("w:pPr")[0];
  if (!pPr) {
    pPr = pNode.ownerDocument.createElement("w:pPr");
    pNode.insertBefore(pPr, pNode.firstChild);
  }
  let spacing = pPr.getElementsByTagName("w:spacing")[0];
  if (!spacing) {
    spacing = pNode.ownerDocument.createElement("w:spacing");
    pPr.appendChild(spacing);
  }
  spacing.setAttribute("w:before", "0");
  spacing.setAttribute("w:after", "0");
  spacing.setAttribute("w:line", "259");
  spacing.setAttribute("w:lineRule", "auto");
}

function updateNodeText(node, regex, replacement) {
  const textNodes = node.getElementsByTagName("w:t");
  for (let i = 0; i < textNodes.length; i++) {
    if (regex.test(textNodes[i].textContent)) {
      textNodes[i].textContent = textNodes[i].textContent.replace(regex, replacement);
    }
  }
}

// Sinh XML bảng tiêu đề chuẩn 2 cột + thông tin thí sinh (Chuẩn Bộ GD&ĐT như ảnh)
function generateExamHeaderXml(info, variantCode) {
  const dept = escapeXml((info.department || '').trim().toUpperCase());
  const school = escapeXml((info.school || '').trim().toUpperCase());
  const exam = escapeXml((info.examName || '').trim().toUpperCase());
  const year = escapeXml((info.schoolYear || '').trim().toUpperCase());
  const subject = escapeXml((info.subject || '').trim().toUpperCase());
  const duration = escapeXml((info.duration || '').trim());
  const code = escapeXml(variantCode);

  return `
  <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:tblPr>
      <w:tblW w:w="10205" w:type="dxa"/>
      <w:jc w:val="center"/>
      <w:tblBorders>
        <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="4600"/>
      <w:gridCol w:w="5605"/>
    </w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="4600" w:type="dxa"/>
          <w:vAlign w:val="top"/>
        </w:tcPr>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:b/>
              <w:color w:val="C00000"/>
              <w:sz w:val="22"/>
            </w:rPr>
            <w:t>${dept}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:b/>
              <w:color w:val="C00000"/>
              <w:sz w:val="22"/>
            </w:rPr>
            <w:t>${school}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="50" w:after="40" w:line="240" w:lineRule="auto"/>
            <w:ind w:left="1175" w:right="1175"/>
            <w:pBdr>
              <w:top w:val="single" w:sz="12" w:space="2" w:color="002060"/>
              <w:left w:val="single" w:sz="12" w:space="6" w:color="002060"/>
              <w:bottom w:val="single" w:sz="12" w:space="2" w:color="002060"/>
              <w:right w:val="single" w:sz="12" w:space="6" w:color="002060"/>
            </w:pBdr>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:b/>
              <w:color w:val="C00000"/>
              <w:sz w:val="22"/>
            </w:rPr>
            <w:t>ĐỀ CHÍNH THỨC</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="20" w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:i/>
              <w:sz w:val="20"/>
            </w:rPr>
            <w:t>(Đề thi có </w:t>
          </w:r>
          <w:fldSimple w:instr="NUMPAGES">
            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                <w:i/>
                <w:sz w:val="20"/>
              </w:rPr>
              <w:t>04</w:t>
            </w:r>
          </w:fldSimple>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:i/>
              <w:sz w:val="20"/>
            </w:rPr>
            <w:t> trang)</w:t>
          </w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="5605" w:type="dxa"/>
          <w:vAlign w:val="top"/>
        </w:tcPr>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:b/>
              <w:sz w:val="22"/>
            </w:rPr>
            <w:t>${exam}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:b/>
              <w:sz w:val="22"/>
            </w:rPr>
            <w:t>${year}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="20" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:b/>
              <w:sz w:val="22"/>
            </w:rPr>
            <w:t>MÔN: ${subject}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:i/>
              <w:sz w:val="20"/>
            </w:rPr>
            <w:t>Thời gian làm bài: ${duration}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="40" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
              <w:i/>
              <w:sz w:val="20"/>
            </w:rPr>
            <w:t>(không kể thời gian phát đề)</w:t>
          </w:r>
        </w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
  <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:pPr>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="6" w:space="8" w:color="000000"/>
      </w:pBdr>
      <w:spacing w:before="120" w:after="160" w:line="260" w:lineRule="auto"/>
      <w:tabs>
        <w:tab w:val="left" w:pos="6500"/>
        <w:tab w:val="right" w:pos="10205"/>
      </w:tabs>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="22"/>
      </w:rPr>
      <w:t xml:space="preserve">Họ và tên: ............................................................................   </w:t>
    </w:r>
    <w:r><w:tab/></w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="22"/>
      </w:rPr>
      <w:t xml:space="preserve">Số báo danh: .......   </w:t>
    </w:r>
    <w:r><w:tab/></w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:b/>
        <w:sz w:val="24"/>
      </w:rPr>
      <w:t>Mã đề ${code}</w:t>
    </w:r>
  </w:p>
  `;
}

// Sinh đoạn kết thúc đề thi: "------ HẾT ------" căn giữa
function generateExamEndXml() {
  return `
  <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:b/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
      <w:t xml:space="preserve">------ HẾT ------</w:t>
    </w:r>
  </w:p>
  `;
}

// Cấu hình Footer chuẩn Word: gạch ngang bên trên, Mã đề bên trái, Trang 1/4 bên phải
async function setupExamFooter(newZip, variantCode) {
  const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pBdr>
        <w:top w:val="single" w:sz="6" w:space="4" w:color="808080"/>
      </w:pBdr>
      <w:tabs>
        <w:tab w:val="right" w:pos="10205"/>
      </w:tabs>
      <w:spacing w:before="60" w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="20"/>
        <w:color w:val="333333"/>
      </w:rPr>
      <w:t>Mã đề ${escapeXml(variantCode)}</w:t>
    </w:r>
    <w:r>
      <w:tab/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="20"/>
        <w:color w:val="333333"/>
      </w:rPr>
      <w:t>Trang </w:t>
    </w:r>
    <w:fldSimple w:instr="PAGE">
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          <w:sz w:val="20"/>
          <w:color w:val="333333"/>
        </w:rPr>
        <w:t>1</w:t>
      </w:r>
    </w:fldSimple>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="20"/>
        <w:color w:val="333333"/>
      </w:rPr>
      <w:t>/</w:t>
    </w:r>
    <w:fldSimple w:instr="NUMPAGES">
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          <w:sz w:val="20"/>
          <w:color w:val="333333"/>
        </w:rPr>
        <w:t>4</w:t>
      </w:r>
    </w:fldSimple>
  </w:p>
</w:ftr>`;

  newZip.file("word/footer1.xml", footerXml);

  const contentTypesFile = newZip.file("[Content_Types].xml");
  if (contentTypesFile) {
    let ctXml = await contentTypesFile.async("text");
    if (!ctXml.includes('PartName="/word/footer1.xml"')) {
      ctXml = ctXml.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
      newZip.file("[Content_Types].xml", ctXml);
    }
  }

  const relsFile = newZip.file("word/_rels/document.xml.rels");
  if (relsFile) {
    let relsXml = await relsFile.async("text");
    if (!relsXml.includes('Target="footer1.xml"')) {
      relsXml = relsXml.replace('</Relationships>', '<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>');
      newZip.file("word/_rels/document.xml.rels", relsXml);
    }
  }
}

// Bắt đầu xử lý trộn đề
async function startProcessing() {
  if (!uploadedFile) {
    alert("Vui lòng chọn file đề thi Word (.docx) trước!");
    return;
  }

  const btnProcess = document.getElementById('btnProcess');
  btnProcess.disabled = true;
  btnProcess.innerHTML = `
    <svg class="animate-spin -ml-1 mr-2 h-4 w-4 text-white" fill="none" viewBox="0 0 24 24">
      <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
      <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
    </svg>
    <span>Đang phân tích &amp; tạo đề...</span>
  `;

  try {
    const useCustomHeader = document.getElementById('chkUseCustomHeader') ? document.getElementById('chkUseCustomHeader').checked : true;
    const commonInfo = {
      department: document.getElementById('infoDepartment') ? document.getElementById('infoDepartment').value : 'SỞ GD&ĐT TP CẦN THƠ',
      school: document.getElementById('infoSchool') ? document.getElementById('infoSchool').value : 'TRƯỜNG TIỂU HỌC, THCS VÀ THPT QUỐC TẾ HÒA BÌNH',
      examName: document.getElementById('infoExamName') ? document.getElementById('infoExamName').value : 'KIỂM TRA GIỮA KỲ I',
      schoolYear: document.getElementById('infoSchoolYear') ? document.getElementById('infoSchoolYear').value : 'NĂM HỌC 2026 – 2027',
      subject: document.getElementById('infoSubject') ? document.getElementById('infoSubject').value : 'TOÁN',
      duration: document.getElementById('infoDuration') ? document.getElementById('infoDuration').value : '50 phút'
    };

    const arrayBuffer = await uploadedFile.arrayBuffer();
    const baseZip = await JSZip.loadAsync(arrayBuffer);
    const docXmlText = await baseZip.file("word/document.xml").async("text");

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(docXmlText, "application/xml");
    const body = xmlDoc.getElementsByTagName("w:body")[0];
    const children = Array.from(body.childNodes);

    const headerNodes = [];
    const sections = [];
    let currentSection = null;
    let currentQuestion = null;
    let inExamBody = false;

    // Phân tích các phần theo thẻ nhóm <#g1> - <#g4> hoặc tiêu đề "Phần I", "Phần II"...
    for (let node of children) {
      if (node.nodeName === 'w:sectPr') continue;

      const textContent = node.textContent || '';

      // Bỏ qua các dòng kết thúc cũ trong file đề gốc (như "------ HẾT ------" hoặc "HẾT") để không bị trùng lặp
      if (/^[\s\-_–—*]*HẾT[\s\-_–—*]*$/i.test(textContent.trim())) {
        continue;
      }

      const detectedType = detectSectionType(node);

      if (detectedType) {
        inExamBody = true;
        cleanGroupTag(node);
        currentSection = {
          type: detectedType,
          titleText: textContent.replace(/<#g\s*[0-4]\s*>\s*/gi, '').trim(),
          titleNode: node.cloneNode(true),
          questions: []
        };
        sections.push(currentSection);
        currentQuestion = null;
        continue;
      }

      const isQuestionStart = /^(Câu|Question)\s+\d+[\.:]/i.test(textContent.trim());

      // Nếu gặp câu hỏi mà chưa vào section nào, tự động tạo Phần I mặc định
      if (isQuestionStart && !currentSection) {
        inExamBody = true;
        currentSection = {
          type: 'g1',
          titleText: 'PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn',
          titleNode: null,
          questions: []
        };
        sections.push(currentSection);
      }

      if (!inExamBody) {
        headerNodes.push(node.cloneNode(true));
        continue;
      }

      if (isQuestionStart) {
        currentQuestion = {
          nodes: [node.cloneNode(true)]
        };
        if (currentSection) currentSection.questions.push(currentQuestion);
      } else if (currentQuestion) {
        currentQuestion.nodes.push(node.cloneNode(true));
      } else if (currentSection) {
        cleanGroupTag(node);
        currentSection.titleNode = node.cloneNode(true);
      }
    }

    const numVariants = parseInt(document.getElementById('numVariants').value) || 4;
    const startCode = parseInt(document.getElementById('startCode').value) || 101;
    generatedBlobs = [];

    // Danh sách mã đề gồm Đề gốc "000" ở đầu, sau đó là 101, 102...
    const variantCodes = ["000"];
    for (let v = 0; v < numVariants; v++) {
      variantCodes.push(String(startCode + v));
    }



    // Cấu trúc lưu trữ đáp án ma trận
    let sectionAnswers = {};
    sections.forEach((sec, sIdx) => {
      let defaultTitle = 'PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn';
      if (sec.type === 'g2') defaultTitle = 'PHẦN II. Trắc nghiệm lựa chọn đúng sai.';
      if (sec.type === 'g3') defaultTitle = 'PHẦN III. Câu hỏi trắc nghiệm trả lời ngắn.';

      sectionAnswers[sIdx] = {
        title: sec.titleText || defaultTitle,
        type: sec.type,
        headers: [],
        rows: {}
      };
    });

    for (let variantCode of variantCodes) {
      const isOriginal = (variantCode === "000");
      const newZip = await JSZip.loadAsync(arrayBuffer);
      const newDoc = parser.parseFromString(docXmlText, "application/xml");
      const newBody = newDoc.getElementsByTagName("w:body")[0];

      while (newBody.firstChild) newBody.removeChild(newBody.firstChild);

      // Thêm Header (Nếu bật tiêu đề chuẩn thì sinh theo thông tin chung, ngược lại dùng header gốc)
      if (useCustomHeader) {
        const headerXml = generateExamHeaderXml(commonInfo, variantCode);
        const headerFragment = parser.parseFromString(
          `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${headerXml}</root>`,
          "application/xml"
        );
        Array.from(headerFragment.documentElement.childNodes).forEach(node => {
          newBody.appendChild(node);
        });
      } else {
        // Thêm Header gốc (làm sạch section break cũ để không làm lệch lề)
        headerNodes.forEach(node => {
          const clone = node.cloneNode(true);
          removeSectionBreaks(clone);
          updateNodeText(clone, /Mã đề.*?:?\s*\d*/i, `Mã đề ${variantCode}`);
          newBody.appendChild(clone);
        });
      }

      // Thêm các phần theo quy tắc nhóm
      sections.forEach((sec, sIdx) => {
        if (sec.titleNode) {
          const clonedTitle = sec.titleNode.cloneNode(true);
          removeSectionBreaks(clonedTitle);
          newBody.appendChild(clonedTitle);
        }

        let questionList = [...sec.questions];

        if (!isOriginal) {
          if (sec.type === 'g1' || sec.type === 'g3') {
            questionList = shuffleArray(questionList);
          }
        }

        const shouldShuffleChoices = !isOriginal && (sec.type === 'g1' || sec.type === 'g2');
        let variantPartAnswers = [];

        questionList.forEach((q, idx) => {
          const newQIndex = idx + 1;
          const parsedMC = parseMultipleChoiceQuestion(q.nodes);
          const tfItems = extractSubItemsTrueFalse(q.nodes);
          const shortAns = extractShortAnswer(q.nodes, sec.type === 'g3');

          // Nhận diện loại câu hỏi: Ưu tiên phân nhóm của phần thi (<#g1>, <#g2>, <#g3>, <#g4>)
          let qType = sec.type;
          if (sec.type === 'g3') {
            qType = 'g3';
          } else if (sec.type === 'g2') {
            qType = 'g2';
          } else if (sec.type === 'g1') {
            qType = 'g1';
          } else if (sec.type === 'g4') {
            qType = 'g4';
          } else {
            if (parsedMC) {
              qType = 'g1';
            } else if (tfItems && tfItems.length === 4) {
              qType = 'g2';
            } else if (shortAns) {
              qType = 'g3';
            }
          }

          if (qType === 'g1' && parsedMC) {
            // === PHẦN I (<#g1>): TRẮC NGHIỆM 4 PHƯƠNG ÁN ===
            // 1. Thêm thân câu hỏi (kể cả bảng biểu, hình vẽ nếu có)
            parsedMC.stemNodes.forEach((sNode, sIdx) => {
              const clonedStem = sNode.cloneNode(true);
              removeSectionBreaks(clonedStem);
              if (sIdx === 0) renumberQuestion(clonedStem, newQIndex);
              applySpacing(clonedStem);
              newBody.appendChild(clonedStem);
            });

            // 2. Dàn cột A, B, C, D (4 cột, 2 cột hoặc 1 cột chuẩn xác)
            const { xmlString, correctChoiceLetter } = generateChoicesXml(parsedMC.choices, shouldShuffleChoices);
            variantPartAnswers.push(correctChoiceLetter || '-');

            if (isOriginal && sectionAnswers[sIdx]) {
              sectionAnswers[sIdx].headers.push(String(newQIndex));
            }

            const choicesFragment = parser.parseFromString(
              `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">${xmlString}</root>`,
              "application/xml"
            );
            Array.from(choicesFragment.documentElement.childNodes).forEach(node => {
              newBody.appendChild(node);
            });
          } else if (qType === 'g2' && tfItems && tfItems.length === 4) {
            // === PHẦN II (<#g2>): ĐÚNG / SAI ===
            q.nodes.forEach((qNode, nodeIdx) => {
              const clonedNode = qNode.cloneNode(true);
              removeSectionBreaks(clonedNode);
              if (nodeIdx === 0) renumberQuestion(clonedNode, newQIndex);
              boldSubQuestions(clonedNode);
              removeUnderline(clonedNode);
              applySpacing(clonedNode);
              newBody.appendChild(clonedNode);
            });

            tfItems.forEach(item => {
              variantPartAnswers.push(item.isCorrect ? 'Đ' : 'S');
              if (isOriginal && sectionAnswers[sIdx]) {
                sectionAnswers[sIdx].headers.push(`${newQIndex}${item.label}`);
              }
            });
          } else if (qType === 'g3') {
            // === PHẦN III (<#g3>): TRẢ LỜI NGẮN (A. [đáp án]) ===
            // Đọc hết tất cả các ký tự bên phải sau ký tự A. làm đáp án
            variantPartAnswers.push(shortAns ? shortAns.answerText : '-');
            if (isOriginal && sectionAnswers[sIdx]) {
              sectionAnswers[sIdx].headers.push(String(newQIndex));
            }

            // Ẩn đáp án A. trong đề thi xuất ra cho học sinh
            q.nodes.forEach((qNode, nodeIdx) => {
              if (shortAns && !shortAns.isInline && nodeIdx === shortAns.nodeIndex) {
                // Đáp án A. nằm trên dòng riêng: ẩn toàn bộ dòng này
                return;
              }

              const clonedNode = qNode.cloneNode(true);
              removeSectionBreaks(clonedNode);
              if (nodeIdx === 0) renumberQuestion(clonedNode, newQIndex);

              // Nếu đáp án A. nằm chung dòng với câu hỏi, xóa bỏ phần A. [đáp án]
              if (shortAns && shortAns.isInline && nodeIdx === shortAns.nodeIndex) {
                removeShortAnswerFromNode(clonedNode);
              }

              applySpacing(clonedNode);
              newBody.appendChild(clonedNode);
            });
          } else {
            // === PHẦN TỰ LUẬN (<#g4>) HOẶC NỘI DUNG CỐ ĐỊNH ===
            q.nodes.forEach((qNode, nodeIdx) => {
              const clonedNode = qNode.cloneNode(true);
              removeSectionBreaks(clonedNode);
              if (nodeIdx === 0) renumberQuestion(clonedNode, newQIndex);
              boldSubQuestions(clonedNode);
              applySpacing(clonedNode);
              newBody.appendChild(clonedNode);
            });
          }
        });

        if (sectionAnswers[sIdx]) {
          sectionAnswers[sIdx].rows[variantCode] = variantPartAnswers;
        }
      });

      // Thêm dòng kết thúc đề thi: "------ HẾT ------" căn giữa
      const endXml = generateExamEndXml();
      const endDoc = parser.parseFromString(
        `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${endXml}</root>`,
        "application/xml"
      );
      Array.from(endDoc.documentElement.childNodes).forEach(node => {
        newBody.appendChild(node);
      });

      // Cấu hình Footer chuẩn theo yêu cầu (gạch ngang bên trên, Mã đề bên trái, Trang 1/4 bên phải)
      await setupExamFooter(newZip, variantCode);

      // Đảm bảo namespace xmlns:r trên document
      if (!newDoc.documentElement.getAttribute("xmlns:r")) {
        newDoc.documentElement.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
      }

      // THIẾT LẬP CHUẨN KHỔ GIẤY A4 VÀ CĂN LỀ:
      // Giấy A4 (Width 21cm = 11906 dxa, Height 29.7cm = 16838 dxa)
      // Căn lề: Top 1cm (567 dxa), Bottom 1cm (567 dxa), Left 2cm (1134 dxa), Right 1cm (567 dxa), Gutter 0cm
      const sectDoc = parser.parseFromString(
        `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:footerReference w:type="default" r:id="rIdFooter1"/>
          <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
          <w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="1134" w:header="720" w:footer="567" w:gutter="0"/>
          <w:cols w:space="720"/>
          <w:docGrid w:linePitch="360"/>
        </w:sectPr>`,
        "application/xml"
      );
      const sectNode = newDoc.importNode(sectDoc.documentElement, true);
      newBody.appendChild(sectNode);

      const serializer = new XMLSerializer();
      newZip.file("word/document.xml", serializer.serializeToString(newDoc));

      const blob = await newZip.generateAsync({ type: "blob" });
      generatedBlobs.push({
        code: variantCode,
        isOriginal: isOriginal,
        blob: blob
      });
    }

    // Tự sinh file Excel bằng OpenXML thuần
    answerExcelBlob = await generateNativeExcelBlob(sectionAnswers, variantCodes);

    displayResults(sections);
  } catch (err) {
    console.error(err);
    alert("Có lỗi trong quá trình xử lý: " + err.message);
  } finally {
    btnProcess.disabled = false;
    btnProcess.innerHTML = `
      <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.2" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z"></path>
      </svg>
      <span>Tiến hành phân tích &amp; tạo đề</span>
    `;
  }
}

// BỘ TẠO FILE EXCEL (.xlsx) OPENXML THUẦN
async function generateNativeExcelBlob(sectionAnswers, variantCodes) {
  try {
    const zip = new JSZip();

    let rowsXml = '';
    let rowIdx = 1;

    Object.values(sectionAnswers).forEach(sec => {
      if (!sec.headers || sec.headers.length === 0) return;

      // 1. Dòng Tiêu đề phần
      rowsXml += `<row r="${rowIdx}">
        <c r="A${rowIdx}" t="inlineStr"><is><t>${escapeXml(sec.title)}</t></is></c>
      </row>`;
      rowIdx++;

      // 2. Dòng tiêu đề: Đề\câu, 1, 2, 3...
      let headerCells = `<c r="A${rowIdx}" t="inlineStr"><is><t>Đề\\câu</t></is></c>`;
      sec.headers.forEach((h, i) => {
        const colLetter = getExcelColName(i + 1);
        headerCells += `<c r="${colLetter}${rowIdx}" t="inlineStr"><is><t>${escapeXml(h)}</t></is></c>`;
      });
      rowsXml += `<row r="${rowIdx}">${headerCells}</row>`;
      rowIdx++;

      // 3. Các dòng mã đề 000, 101, 102...
      variantCodes.forEach(code => {
        let dataCells = `<c r="A${rowIdx}" t="inlineStr"><is><t>${escapeXml(code)}</t></is></c>`;
        const answers = sec.rows[code] || [];
        answers.forEach((ans, i) => {
          const colLetter = getExcelColName(i + 1);
          dataCells += `<c r="${colLetter}${rowIdx}" t="inlineStr"><is><t>${escapeXml(ans)}</t></is></c>`;
        });
        rowsXml += `<row r="${rowIdx}">${dataCells}</row>`;
        rowIdx++;
      });

      rowIdx++;
    });

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>${rowsXml}</sheetData>
    </worksheet>`;

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Đáp án" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`;

    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    </Relationships>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    </Types>`;

    zip.file("[Content_Types].xml", contentTypesXml);
    zip.folder("_rels").file(".rels", relsXml);
    zip.folder("xl").file("workbook.xml", workbookXml);
    zip.folder("xl").folder("_rels").file("workbook.xml.rels", workbookRelsXml);
    zip.folder("xl").folder("worksheets").file("sheet1.xml", sheetXml);

    return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  } catch (e) {
    console.error("Lỗi tạo Excel OpenXML:", e);
    return null;
  }
}

function getExcelColName(colIndex) {
  let colName = '';
  while (colIndex >= 0) {
    colName = String.fromCharCode((colIndex % 26) + 65) + colName;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return colName;
}

function escapeXml(unsafe) {
  if (!unsafe) return '';
  return String(unsafe).replace(/[<>&'"]/g, c => {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
    }
  });
}

function displayResults(sections) {
  const totalQ = sections.reduce((sum, s) => sum + s.questions.length, 0);
  const statsInfo = document.getElementById('statsInfo');
  if (statsInfo) {
    statsInfo.innerText = `Đã tạo Đề gốc 000, các mã đề và file Excel đáp án (${totalQ} câu hỏi).`;
  }

  const variantList = document.getElementById('variantList');
  if (!variantList) return;
  variantList.innerHTML = '';

  // Hiển thị các mã đề Word
  generatedBlobs.forEach(v => {
    const item = document.createElement('div');
    const isOrig = v.isOriginal;
    item.className = `variant-card p-3 rounded-xl border flex flex-col justify-between items-center text-center space-y-2.5 transition shadow-2xs hover:shadow-md ${isOrig ? 'bg-gradient-to-b from-amber-50 to-orange-50/40 border-amber-200' : 'bg-white border-slate-200/90'
      }`;
    item.innerHTML = `
      <div class="w-full flex items-center justify-between pb-1 border-b ${isOrig ? 'border-amber-200/60' : 'border-slate-100'}">
        <span class="text-[12px] font-mono font-bold px-1.5 py-0.5 rounded ${isOrig ? 'bg-amber-200/70 text-amber-800' : 'bg-blue-50 text-blue-700'}">
          ${isOrig ? 'GỐC' : 'HOÁN VỊ'}
        </span>
        <span class="text-[12px] text-slate-400 font-semibold font-mono">.DOCX</span>
      </div>
      <div class="py-1">
        <div class="text-xs sm:text-sm font-extrabold ${isOrig ? 'text-amber-900' : 'text-slate-800'}">
          ${isOrig ? 'Đề gốc (000)' : `Mã đề ${v.code}`}
        </div>
      </div>
      <button class="w-full text-xs py-1.5 px-2 rounded-lg font-bold text-white transition flex items-center justify-center space-x-1.5 shadow-xs ${isOrig ? 'bg-amber-600 hover:bg-amber-700 shadow-amber-600/20' : 'bg-blue-600 hover:bg-blue-700 shadow-blue-600/20'
      }">
        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
        <span>Tải Word</span>
      </button>
    `;
    item.querySelector('button').addEventListener('click', () => {
      const fileName = isOrig ? "De_Goc_Ma_000.docx" : `De_Thi_Ma_${v.code}.docx`;
      saveAs(v.blob, fileName);
    });
    variantList.appendChild(item);
  });

  // Thêm thẻ tải riêng file Excel ma trận đáp án nếu có
  if (answerExcelBlob) {
    const excelItem = document.createElement('div');
    excelItem.className = `variant-card p-3 rounded-xl border border-emerald-200 bg-gradient-to-b from-emerald-50 to-teal-50/40 flex flex-col justify-between items-center text-center space-y-2.5 transition shadow-2xs hover:shadow-md`;
    excelItem.innerHTML = `
      <div class="w-full flex items-center justify-between pb-1 border-b border-emerald-200/60">
        <span class="text-[12px] font-mono font-bold px-1.5 py-0.5 rounded bg-emerald-200/70 text-emerald-800">
          ĐÁP ÁN
        </span>
        <span class="text-[12px] text-emerald-600 font-semibold font-mono">.XLSX</span>
      </div>
      <div class="py-1">
        <div class="text-xs sm:text-sm font-extrabold text-emerald-900">
          Ma trận đáp án
        </div>
      </div>
      <button class="w-full text-xs py-1.5 px-2 rounded-lg font-bold text-white transition flex items-center justify-center space-x-1.5 shadow-xs bg-emerald-600 hover:bg-emerald-700 shadow-emerald-600/20">
        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
        <span>Tải Excel</span>
      </button>
    `;
    excelItem.querySelector('button').addEventListener('click', () => {
      saveAs(answerExcelBlob, "Dap_An.xlsx");
    });
    variantList.appendChild(excelItem);
  }

  const resultCard = document.getElementById('resultCard');
  if (resultCard) resultCard.classList.remove('hidden');
}

// Tải file ZIP trọn gói
async function downloadAllZip() {
  if (!generatedBlobs || generatedBlobs.length === 0) return;

  const zip = new JSZip();

  generatedBlobs.forEach(v => {
    const fileName = v.isOriginal ? "De_Goc_Ma_000.docx" : `De_Thi_Ma_${v.code}.docx`;
    zip.file(fileName, v.blob);
  });

  if (answerExcelBlob) {
    zip.file("Dap_An.xlsx", answerExcelBlob);
  }

  const content = await zip.generateAsync({ type: "blob" });
  saveAs(content, "smartmix.zip");
}

// ==========================================================================
// GOOGLE IDENTITY & AUTHENTICATION MODULE
// ==========================================================================

function initGoogleAuth() {
  const btnGoogleLogin = document.getElementById('btnGoogleLogin');
  const userProfileBox = document.getElementById('userProfileBox');
  const userProfileBtn = document.getElementById('userProfileBtn');
  const userMenuDropdown = document.getElementById('userMenuDropdown');
  const userAvatarImg = document.getElementById('userAvatarImg');
  const userNameLabel = document.getElementById('userNameLabel');
  const menuAvatarImg = document.getElementById('menuAvatarImg');
  const menuUserName = document.getElementById('menuUserName');
  const menuUserEmail = document.getElementById('menuUserEmail');
  const btnGoogleLogout = document.getElementById('btnGoogleLogout');
  
  const googleAuthModal = document.getElementById('googleAuthModal');
  const btnCloseAuthModal = document.getElementById('btnCloseAuthModal');
  const btnConfirmQuickLogin = document.getElementById('btnConfirmQuickLogin');
  const inputDemoName = document.getElementById('inputDemoName');
  const inputDemoEmail = document.getElementById('inputDemoEmail');

  const authToast = document.getElementById('authToast');
  const authToastMsg = document.getElementById('authToastMsg');

  let toastTimer = null;
  function showToast(message) {
    if (!authToast || !authToastMsg) return;
    authToastMsg.textContent = message;
    authToast.classList.remove('hidden');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      authToast.classList.add('hidden');
    }, 3500);
  }

  function generateAvatarDataUrl(name) {
    const initial = (name ? name.trim().charAt(0) : 'U').toUpperCase();
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64"><rect width="64" height="64" rx="32" fill="#2563eb"/><text x="50%" y="54%" dominant-baseline="middle" text-anchor="middle" fill="#ffffff" font-family="sans-serif" font-size="28" font-weight="bold">${initial}</text></svg>`;
    return 'data:image/svg+xml;utf8,' + encodeURIComponent(svg);
  }

  function renderUserUI(user) {
    if (user) {
      if (btnGoogleLogin) btnGoogleLogin.classList.add('hidden');
      if (userProfileBox) userProfileBox.classList.remove('hidden');

      const avatarSrc = user.picture || generateAvatarDataUrl(user.name);
      if (userAvatarImg) userAvatarImg.src = avatarSrc;
      if (menuAvatarImg) menuAvatarImg.src = avatarSrc;
      if (userNameLabel) userNameLabel.textContent = user.name || 'Người dùng';
      if (menuUserName) menuUserName.textContent = user.name || 'Người dùng';
      if (menuUserEmail) menuUserEmail.textContent = user.email || '';
    } else {
      if (btnGoogleLogin) btnGoogleLogin.classList.remove('hidden');
      if (userProfileBox) userProfileBox.classList.add('hidden');
      if (userMenuDropdown) userMenuDropdown.classList.add('hidden');
    }
  }

  function openAuthModal() {
    if (!googleAuthModal) return;
    googleAuthModal.classList.remove('hidden');
  }

  function closeAuthModal() {
    if (!googleAuthModal) return;
    googleAuthModal.classList.add('hidden');
  }

  function loginUser(userData) {
    localStorage.setItem('smartmix_user', JSON.stringify(userData));
    renderUserUI(userData);
    closeAuthModal();
    showToast(`Chào mừng ${userData.name}, bạn đã đăng nhập thành công!`);
  }

  function logoutUser() {
    localStorage.removeItem('smartmix_user');
    renderUserUI(null);
    showToast('Đã đăng xuất khỏi tài khoản.');
  }

  // Gán sự kiện
  if (btnGoogleLogin) {
    btnGoogleLogin.addEventListener('click', openAuthModal);
  }

  if (btnCloseAuthModal) {
    btnCloseAuthModal.addEventListener('click', closeAuthModal);
  }

  if (googleAuthModal) {
    googleAuthModal.addEventListener('click', (e) => {
      if (e.target === googleAuthModal) closeAuthModal();
    });
  }

  if (userProfileBtn) {
    userProfileBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (userMenuDropdown) {
        userMenuDropdown.classList.toggle('hidden');
      }
    });
  }

  document.addEventListener('click', (e) => {
    if (userProfileBox && !userProfileBox.contains(e.target)) {
      if (userMenuDropdown) userMenuDropdown.classList.add('hidden');
    }
  });

  if (btnGoogleLogout) {
    btnGoogleLogout.addEventListener('click', logoutUser);
  }

  if (btnConfirmQuickLogin) {
    btnConfirmQuickLogin.addEventListener('click', () => {
      const name = (inputDemoName && inputDemoName.value.trim()) || 'Quách Nhị';
      const email = (inputDemoEmail && inputDemoEmail.value.trim()) || 'nhicnttcantho@gmail.com';
      loginUser({
        name: name,
        email: email,
        picture: generateAvatarDataUrl(name),
        provider: 'google'
      });
    });
  }

  // Tải trạng thái đăng nhập từ localStorage khi mở trang
  try {
    const savedUserJson = localStorage.getItem('smartmix_user');
    if (savedUserJson) {
      const savedUser = JSON.parse(savedUserJson);
      renderUserUI(savedUser);
    } else {
      renderUserUI(null);
    }
  } catch (err) {
    renderUserUI(null);
  }
}

// Tự động khởi chạy Google Auth khi DOM sẵn sàng
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initGoogleAuth);
} else {
  initGoogleAuth();
}