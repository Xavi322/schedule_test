import os
import re
import json
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def ensure_template(template_path):
    """Создает template.html с выбором ПОТОКА (Очное/ИУП), НЕДЕЛИ и ГРУППЫ."""
    lines = [
        '<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
        '<meta name="theme-color" content="#1e293b">',
        '<link rel="manifest" href="manifest.json">',
        '<link rel="icon" type="image/png" sizes="192x192" href="app-icon-192.png">',
        '<link rel="icon" type="image/png" sizes="512x512" href="app-icon-512.png">',
        '<link rel="apple-touch-icon" sizes="192x192" href="app-icon-192.png">',
        '<title>Расписание занятий</title><style>',
        ':root{--bg:#f4f6f9;--card:#fff;--pri:#3b82f6;--pri-l:#eff6ff;--pur:#7c3aed;--pur-l:#f5f3ff;--pur-b:#ddd6fe;--warn:#ea580c;--warn-l:#fff7ed;--warn-b:#fed7aa;--dng:#dc2626;--dng-l:#fef2f2;--dng-b:#fecaca;--txt:#1e293b;--mut:#64748b;--brd:#e2e8f0;--grn:#10b981}',
        '*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}',
        'body{background:var(--bg);color:var(--txt);padding:16px}',
        'header{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:16px;padding:12px 16px;background:var(--card);border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.05)}',
        'header h1{font-size:1.2rem;font-weight:700}',
        '.controls{display:flex;align-items:center;gap:8px;flex-wrap:wrap}',
        '.mode-select{background:#475569;color:#fff;padding:8px 12px;border-radius:8px;font-weight:600;font-size:.9rem;border:none;cursor:pointer;outline:none}',
        '.mode-select option{background:#fff;color:var(--txt)}',
        '.week-select{background:#1e293b;color:#fff;padding:8px 12px;border-radius:8px;font-weight:600;font-size:.9rem;border:none;cursor:pointer;outline:none}',
        '.week-select option{background:#fff;color:var(--txt)}',
        '.group-select{background:var(--pri);color:#fff;padding:8px 14px;border-radius:8px;font-weight:600;font-size:.95rem;border:none;cursor:pointer;outline:none}',
        '.group-select option{background:#fff;color:var(--txt)}',
        '.save-btn{background:#e2e8f0;color:var(--txt);padding:8px 12px;border-radius:8px;font-weight:600;font-size:.85rem;border:none;cursor:pointer;transition:.2s}',
        '.save-btn.is-saved{background:var(--grn);color:#fff}',
        '.week-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;align-items:start}',
        '.day-column{background:var(--card);border-radius:12px;border:1px solid var(--brd);overflow:hidden;box-shadow:0 1px 2px rgba(0,0,0,.04)}',
        '.day-header{background:#1e293b;color:#fff;padding:10px 12px;font-weight:600;font-size:.95rem;text-align:center}',
        '.day-slots{padding:8px;display:flex;flex-direction:column;gap:8px}',
        '.slot-card{border-radius:8px;padding:10px;min-height:92px;display:flex;flex-direction:column;justify-content:space-between}',
        '.active-slot{background:var(--pri-l);border:1px solid #bfdbfe;border-left:4px solid var(--pri)}',
        '.moved-slot{background:var(--pur-l);border:1px solid var(--pur-b);border-left:4px solid var(--pur)}',
        '.shifted-slot{background:var(--warn-l);border:1px solid var(--warn-b);border-left:4px solid var(--warn)}',
        '.empty-slot{background:#f8fafc;border:1px dashed #cbd5e1;color:#94a3b8;min-height:60px}',
        '.slot-header{display:flex;justify-content:space-between;align-items:center;font-size:.75rem;font-weight:600;color:var(--mut);margin-bottom:6px}',
        '.old-time{text-decoration:line-through;color:#94a3b8;font-size:.7rem;margin-right:4px}',
        '.custom-time{background:var(--warn);color:#fff;padding:2px 6px;border-radius:4px;font-weight:700;font-size:.75rem}',
        '.subject-title{font-size:.85rem;font-weight:600;line-height:1.35;color:var(--txt);margin-bottom:8px;white-space:pre-line}',
        '.empty-label{font-size:.8rem;text-align:center;margin:auto 0}',
        '.slot-footer{display:flex;justify-content:space-between;align-items:center;gap:4px;flex-wrap:wrap}',
        '.room-badge{background:#fff;border:1px solid #cbd5e1;color:#0f172a;font-size:.75rem;font-weight:600;padding:2px 8px;border-radius:6px;margin-left:auto}',
        '.moved-badge{background:var(--pur);color:#fff;font-size:.7rem;font-weight:600;padding:2px 6px;border-radius:6px}',
        '.online-btn{background:var(--grn);color:#fff;border:none;font-size:.75rem;font-weight:600;padding:4px 8px;border-radius:6px;cursor:pointer}',
        '.day-off{padding:32px 12px;text-align:center;color:var(--mut);font-weight:500}',
        '.cancelled-section{margin:0 8px 8px 8px;padding:8px 10px;background:var(--dng-l);border:1px solid var(--dng-b);border-radius:8px;font-size:.75rem;color:#7f1d1d}',
        '.cancelled-title{font-weight:700;color:var(--dng);margin-bottom:6px}',
        '.cancelled-item{margin-top:5px;line-height:1.35}',
        '.reason-tag{display:inline-block;background:#fff;border:1px solid var(--dng-b);color:var(--dng);font-weight:700;font-size:.68rem;padding:1px 5px;border-radius:4px;margin-right:3px}',
        '.modal-overlay{display:none;position:fixed;inset:0;background:rgba(15,23,42,.6);align-items:center;justify-content:center;z-index:1000;padding:16px}',
        '.modal-box{background:#fff;padding:20px;border-radius:12px;max-width:420px;width:100%;box-shadow:0 10px 25px rgba(0,0,0,.15)}',
        '.modal-title{font-weight:700;font-size:1rem;margin-bottom:10px}',
        '.modal-link{word-break:break-all;background:#f1f5f9;padding:10px;border-radius:8px;font-size:.85rem;margin-bottom:14px;border:1px solid var(--brd)}',
        '.modal-link a{color:var(--pri);font-weight:600;text-decoration:none}',
        '.modal-actions{display:flex;gap:8px;justify-content:flex-end}',
        '.btn{padding:8px 14px;border-radius:8px;font-weight:600;font-size:.85rem;cursor:pointer;border:none}',
        '.btn-copy{background:var(--pri);color:#fff}',
        '.btn-close{background:#e2e8f0;color:var(--txt)}',
        '</style></head><body>',
        '<header><h1>📅 Расписание</h1>',
        '<div class="controls">',
        '<select id="modeSelect" class="mode-select" autocomplete="off"></select>',
        '<select id="weekSelect" class="week-select" autocomplete="off"></select>',
        '<select id="groupSelect" class="group-select" autocomplete="off"></select>',
        '<button id="saveGroupBtn" class="save-btn" onclick="saveMyGroup()">📌 Запомнить</button>',
        '</div></header>',
        '<div id="weekGrid" class="week-grid"></div>',
        '<div id="linkModal" class="modal-overlay" onclick="if(event.target===this)closeModal()">',
        '<div class="modal-box"><div class="modal-title">💻 Ссылка на подключение</div>',
        '<div class="modal-link"><a id="modalLinkAnchor" href="#" target="_blank"></a></div>',
        '<div class="modal-actions">',
        '<button class="btn btn-copy" onclick="copyModalLink()">Скопировать</button>',
        '<button class="btn btn-close" onclick="closeModal()">Закрыть</button>',
        '</div></div></div><script>',
        'var appData = __APP_DATA_JSON__;',
        'function showLinkModal(u){var a=document.getElementById("modalLinkAnchor");a.href=u;a.textContent=u;document.getElementById("linkModal").style.display="flex";}',
        'function closeModal(){document.getElementById("linkModal").style.display="none";}',
        'function copyModalLink(){navigator.clipboard.writeText(document.getElementById("modalLinkAnchor").textContent);alert("Ссылка скопирована!");}',
        'var modeEl=document.getElementById("modeSelect");',
        'var weekEl=document.getElementById("weekSelect");',
        'var selectEl=document.getElementById("groupSelect");',
        'var saveBtn=document.getElementById("saveGroupBtn");',
        'var grid=document.getElementById("weekGrid");',
        'var STORAGE_KEY="pinned_group_v4";',
        'var modeKeys=Object.keys(appData);',
        'modeKeys.forEach(function(k){var o=document.createElement("option");o.value=k;o.textContent=appData[k].title;modeEl.appendChild(o);});',
        'if(modeKeys.length<=1){modeEl.style.display="none";}',
        'function getSavedState(){',
        '  try{',
        '    var v=localStorage.getItem(STORAGE_KEY);if(v)return v;',
        '    var m=document.cookie.match(/(?:^|; )pinned_group_v4=([^;]*)/);if(m)return decodeURIComponent(m.pop());',
        '    if(window.name)return window.name;',
        '  }catch(e){}',
        '  return "";',
        '}',
        'function populateMode(modeKey,preferredGroup){',
        '  var mObj=appData[modeKey]||appData[modeKeys.slice().shift()];',
        '  var weeksPayload=mObj.weeks||[];',
        '  weekEl.innerHTML="";',
        '  var currIdx=0;',
        '  weeksPayload.forEach(function(w,idx){',
        '    var o=document.createElement("option");o.value=idx;o.textContent=w.label;',
        '    if(w.is_current)currIdx=idx;',
        '    weekEl.appendChild(o);',
        '  });',
        '  weekEl.value=currIdx;',
        '  weekEl.style.display=(weeksPayload.length<=1)?"none":"";',
        '  var firstW=weeksPayload.slice().shift()||{groups:{}};',
        '  var grps=Object.keys(firstW.groups);',
        '  selectEl.innerHTML="";',
        '  grps.forEach(function(g){var o=document.createElement("option");o.value=g;o.textContent=g;selectEl.appendChild(o);});',
        '  if(preferredGroup&&firstW.groups[preferredGroup]){selectEl.value=preferredGroup;}',
        '}',
        'function updateSaveBtn(){',
        '  var saved=getSavedState();',
        '  var cur=modeEl.value+"|"+selectEl.value;',
                'function updateSaveBtn(){',
        '  var saved=getSavedState();',
        '  var cur=modeEl.value+"|"+selectEl.value;',
        '  if(saved&&saved===cur){saveBtn.textContent="✅ Моя группа";saveBtn.className="save-btn is-saved";}',
        '  else{saveBtn.textContent="📌 Запомнить";saveBtn.className="save-btn";}',
        '}',
        'function saveMyGroup(){',
        '  var val=modeEl.value+"|"+selectEl.value;',
        '  try{localStorage.clear();localStorage.setItem(STORAGE_KEY,val);document.cookie="pinned_group_v4="+encodeURIComponent(val)+"; path=/; max-age=31536000";window.name=val;}catch(e){}',
        '  updateSaveBtn();',
        '}',
        'function renderCurrentView(){',
        '  grid.innerHTML="";',
        '  var mObj=appData[modeEl.value];if(!mObj)return;',
        '  var wIdx=parseInt(weekEl.value,10)||0;',
        '  var weekObj=mObj.weeks[wIdx]||mObj.weeks.slice().shift()||{groups:{}};',
        '  var groupName=selectEl.value;',
        '  var scheduleData=(weekObj.groups&&weekObj.groups[groupName])||[];',
        '  updateSaveBtn();',
        '  scheduleData.forEach(function(d){',
        '    var col=document.createElement("div");col.className="day-column";',
        '    var h=document.createElement("div");h.className="day-header";h.textContent=d.day;col.appendChild(h);',
        '    var sb=document.createElement("div");sb.className="day-slots";',
        '    var has=d.slots.some(function(s){return s.subject;});',
        '    if(!has){sb.innerHTML=\'<div class="day-off">🎉 Выходной</div>\';}else{',
        '      d.slots.forEach(function(s){',
        '        var c=document.createElement("div");',
        '        if(s.subject){',
        '          c.className="slot-card active-slot"+(s.moved_from?" moved-slot":"")+(s.custom_time?" shifted-slot":"");',
        '          var th=s.custom_time?\'<span class="old-time">\'+s.time+\'</span> <span class="custom-time">⏰ \'+s.custom_time+\'</span>\':\'<span class="pair-time">\'+s.time+\'</span>\';',
        '          c.innerHTML=\'<div class="slot-header"><span class="pair-num">\'+s.num+\' пара</span><div>\'+th+\'</div></div><div class="subject-title"></div><div class="slot-footer"></div>\';',
        '          c.querySelector(".subject-title").textContent=s.subject;',
        '          var f=c.querySelector(".slot-footer");',
        '          if(s.moved_from){var mb=document.createElement("span");mb.className="moved-badge";mb.textContent="🔄 с "+s.moved_from;f.appendChild(mb);}',
        '          if(s.link){let u=s.link;var b=document.createElement("button");b.className="online-btn";b.textContent="💻 Ссылка";b.onclick=function(){showLinkModal(u);};f.appendChild(b);}',
        '          if(s.room){var rb=document.createElement("span");rb.className="room-badge";rb.textContent="📍 "+s.room;f.appendChild(rb);}',
        '        }else{',
        '          c.className="slot-card empty-slot";',
        '          c.innerHTML=\'<div class="slot-header"><span class="pair-num">\'+s.num+\' пара</span><span class="pair-time">\'+s.time+\'</span></div><div class="empty-label">Свободно</div>\';',
        '        }',
        '        sb.appendChild(c);',
        '      });',
        '    }',
        '    col.appendChild(sb);',
                '    if(d.removed&&d.removed.length>0){',
        '      var rs=document.createElement("div");rs.className="cancelled-section";',
        '      rs.innerHTML=\'<div class="cancelled-title">🚫 Снято с расписания:</div>\';',
        '      d.removed.forEach(function(r){',
        '        var it=document.createElement("div");it.className="cancelled-item";',
        '        it.innerHTML=\'<span class="reason-tag">\'+r.reason+\'</span> <b>\'+r.num+\' пара:</b> <s>\'+r.subject+\'</s>\';',
        '        rs.appendChild(it);',
        '      });',
        '      col.appendChild(rs);',
        '    }',
        '    grid.appendChild(col);',
        '  });',
        '}',
        'modeEl.onchange=function(){populateMode(this.value,"");renderCurrentView();};',
        'selectEl.onchange=renderCurrentView;',
        'weekEl.onchange=renderCurrentView;',
        'function forceRestoreGroup(){',
        '  var saved=getSavedState();',
        '  var defMode=modeKeys.slice().shift();',
        '  var targetMode=defMode,targetGroup="";',
        '  if(saved&&saved.indexOf("|")!==-1){',
        '    var parts=saved.split("|");',
        '    var mCandidate=parts.slice().shift();',
        '    var gCandidate=parts.pop();',
        '    if(appData[mCandidate]){targetMode=mCandidate;targetGroup=gCandidate;}',
        '  }',
        '  if(modeEl.value!==targetMode||(targetGroup&&selectEl.value!==targetGroup)||grid.innerHTML===""){',
        '    modeEl.value=targetMode;',
        '    populateMode(targetMode,targetGroup);',
        '    renderCurrentView();',
        '  }',
        '}',
        'populateMode(modeKeys.slice().shift(),"");',
        'forceRestoreGroup();',
        'setTimeout(forceRestoreGroup,150);',
        'setTimeout(forceRestoreGroup,500);',
        'window.addEventListener("pageshow",function(){setTimeout(forceRestoreGroup,50);});',
        'if("serviceWorker" in navigator){navigator.serviceWorker.register("sw.js");}',
        '</script></body></html>'
    ]
    with open(template_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

def generate_week_calendar(filename="schedule.xlsx", links_file="links.xlsx", output_html="index.html"):
    template_path = os.path.join(BASE_DIR, "template.html")
    output_path = os.path.join(BASE_DIR, output_html)
    meta_path = os.path.join(BASE_DIR, "weeks_meta.json")

    ensure_template(template_path)
    app_data = {}

    if os.path.exists(meta_path):
        with open(meta_path, "r", encoding="utf-8") as mf:
            raw_meta = json.load(mf)

        if isinstance(raw_meta, dict):
            for mode_key, mode_info in raw_meta.items():
                l_path = os.path.join(BASE_DIR, mode_info.get("links_file", links_file))
                if not os.path.exists(l_path):
                    l_path = os.path.join(BASE_DIR, links_file)
                mode_links = load_all_online_links(l_path)

                weeks_list = []
                for w in mode_info.get("weeks", []):
                    f_path = os.path.join(BASE_DIR, w["file"])
                    if os.path.exists(f_path):
                        weeks_list.append({
                            "label": w["label"],
                            "is_current": w.get("is_current", False),
                            "groups": parse_workbook_all_groups(f_path, mode_links)
                        })
                if weeks_list:
                    app_data[mode_key] = {"title": mode_info.get("title", mode_key), "weeks": weeks_list}

    if not app_data:
        sched_path = os.path.join(BASE_DIR, filename)
        l_path = os.path.join(BASE_DIR, links_file)
        if os.path.exists(sched_path):
            app_data["main"] = {
                "title": "🎓 Очное",
                "weeks": [{
                    "label": "📅 Текущая неделя",
                    "is_current": True,
                    "groups": parse_workbook_all_groups(sched_path, load_all_online_links(l_path))
                }]
            }

    with open(template_path, "r", encoding="utf-8") as tf:
        html_content = tf.read()

    html_content = html_content.replace("__APP_DATA_JSON__", json.dumps(app_data, ensure_ascii=False))

    with open(output_path, "w", encoding="utf-8") as out:
        out.write(html_content)
    print(f"✅ Собрано потоков в календарь: {len(app_data)} -> {output_path}")

if __name__ == "__main__":
    generate_week_calendar("schedule.xlsx", "links.xlsx", "index.html")


