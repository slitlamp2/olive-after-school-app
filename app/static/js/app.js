/* 올리브청소년방과후 센터 활동앱 – 공통 JS */

/* ── 토스트 ── */
function showToast(msg, type) {
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast' + (type === 'error' ? ' error' : '');
  t.classList.remove('hidden');
  setTimeout(function() { t.classList.add('hidden'); }, 4000);
}

/* ── 로딩 ── */
function showLoading(msg) {
  document.getElementById('loading-msg').textContent = msg || '처리 중입니다…';
  document.getElementById('loading-overlay').classList.remove('hidden');
}
function hideLoading() {
  document.getElementById('loading-overlay').classList.add('hidden');
}

/* ── 사진 미리보기 (단일 multiple 입력) ── */
function setupPhotoPreview(inputId, previewId, maxCount) {
  var input   = document.getElementById(inputId);
  var preview = document.getElementById(previewId);
  if (!input || !preview) return;

  input.addEventListener('change', function() {
    preview.innerHTML = '';
    var files = Array.from(this.files);
    if (maxCount && files.length > maxCount) {
      showToast('최대 ' + maxCount + '장까지 업로드할 수 있습니다.', 'error');
      this.value = '';
      return;
    }
    files.forEach(function(file) {
      var reader = new FileReader();
      reader.onload = function(e) {
        var div = document.createElement('div');
        div.className = 'preview-item';
        div.innerHTML = '<img src="' + e.target.result + '" alt="미리보기">';
        preview.appendChild(div);
      };
      reader.readAsDataURL(file);
    });
  });
}

/* ── 사진 업로드 버튼 (각 버튼당 1장, 미리보기 슬롯) ── */
function setupPhotoButtons(inputIds, previewIds) {
  inputIds.forEach(function(inputId, i) {
    var input   = document.getElementById(inputId);
    var preview = document.getElementById(previewIds[i]);
    if (!input || !preview) return;

    var btn = document.querySelector('.upload-btn[data-for="' + inputId + '"]');
    if (btn) {
      btn.addEventListener('click', function() { input.click(); });
    }

    input.addEventListener('change', function() {
      var file = this.files && this.files[0];
      preview.innerHTML = '';
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function(e) {
        var img = document.createElement('img');
        img.src = e.target.result;
        img.alt = '미리보기';
        preview.appendChild(img);
      };
      reader.readAsDataURL(file);
    });
  });
}

/* ── 폼 제출 (multipart/form-data → JSON 응답) ── */
function setupFormSubmit(formId, endpoint, onSuccess) {
  var form = document.getElementById(formId);
  if (!form) return;

  form.addEventListener('submit', function(e) {
    e.preventDefault();
    var fd = new FormData(form);
    showLoading('AI 분석 및 HWP 생성 중… 잠시 기다려 주세요.');

    var url = (window.location.origin || '') + endpoint;
    fetch(url, { method: 'POST', body: fd })
      .then(function(r) {
        return r.text().then(function(text) {
          var data;
          try {
            data = text ? JSON.parse(text) : {};
          } catch (e) {
            if (!r.ok) {
              throw new Error('서버 오류 ' + r.status + '. 응답을 읽을 수 없습니다.');
            }
            throw new Error('응답 형식 오류');
          }
          if (!r.ok) {
            var errMsg = (data && data.error) ? data.error : ('서버 오류 ' + r.status);
            throw new Error(errMsg);
          }
          return data;
        });
      })
      .then(function(data) {
        hideLoading();
        if (data.success) {
          showToast(data.message || '생성 완료!');
          if (onSuccess) onSuccess(data);
        } else {
          showToast('오류: ' + (data.error || '알 수 없는 오류'), 'error');
        }
      })
      .catch(function(err) {
        hideLoading();
        var msg = (err && err.message) ? err.message : String(err);
        if (msg.indexOf('fetch') !== -1 || msg.indexOf('Failed') !== -1) {
          showToast('서버에 연결할 수 없습니다. 서버가 실행 중인지 확인하고, 주소창이 http://localhost:5006 인지 확인해 주세요.', 'error');
        } else if (msg.indexOf('서버 오류') !== -1 || msg.indexOf('응답') !== -1) {
          showToast(msg, 'error');
        } else {
          showToast('오류: ' + msg, 'error');
        }
      });
  });
}
