(function(){
  const baseConfig={
    pageKey:'user',
    realName:'用户',
    role:'用户',
    account:'user001',
    personId:'',
    organization:'',
    email:'None',
    phone:'',
    defaultAvatarText:'用',
    description:'账号信息与个性化设置。',
    statsTitle:'账号概览',
    stats:[],
    actions:[{label:'修改密码'},{label:'通知设置'},{label:'主题设置'}],
    allowLogout:false
  };
  const config=Object.assign({},baseConfig,window.PROFILE_CONFIG||{});
  const storageKey='ai-course-profile-'+(config.pageKey||config.account);
  const state=loadState();

  function init(){
    loadProfileCss();
    ensureTeacherHeaderName();
    ensureProfileSection();
    renderProfile();
    applyIdentity();
  }

  function loadProfileCss(){
    if(document.querySelector('link[data-profile-settings]'))return;
    const link=document.createElement('link');
    link.rel='stylesheet';
    link.href='profile-settings.css';
    link.dataset.profileSettings='true';
    document.head.appendChild(link);
  }

  function loadState(){
    const defaults={nickname:'',displayMode:'real',avatarData:''};
    try{
      const raw=localStorage.getItem(storageKey);
      return Object.assign(defaults,raw?JSON.parse(raw):{});
    }catch(e){
      return defaults;
    }
  }

  function saveState(){
    try{
      localStorage.setItem(storageKey,JSON.stringify(state));
      return true;
    }catch(e){
      showStatus('本地保存空间不足，请选择更小的图片。','bad');
      return false;
    }
  }

  function ensureProfileSection(){
    let section=document.getElementById('profile');
    if(section)return section;
    const content=document.querySelector('.content');
    section=document.createElement('section');
    section.id='profile';
    section.className='section';
    if(content)content.appendChild(section);
    return section;
  }

  function ensureTeacherHeaderName(){
    const topRight=document.querySelector('.top-right');
    if(!topRight||topRight.querySelector('.user-name'))return;
    const avatar=topRight.querySelector('.avatar');
    if(!avatar)return;
    const name=document.createElement('span');
    name.className='user-name';
    avatar.insertAdjacentElement('afterend',name);
  }

  function renderProfile(){
    const section=ensureProfileSection();
    section.innerHTML=[
      '<div class="profile-page">',
      '<div class="profile-panel">',
      '<div class="profile-block"><div class="profile-block-title">基本信息</div><div class="profile-rows">',
      renderInfoRow('姓名',escapeHtml(config.realName)),
      '<div class="profile-row"><div class="profile-label">昵称</div><div class="profile-value"><div class="profile-fieldline"><input class="profile-input" id="profileNickname" maxlength="20" placeholder="请输入昵称"><button class="profile-btn" id="profileSave" type="button">保存</button></div></div></div>',
      '<div class="profile-row"><div class="profile-label">展示方式</div><div class="profile-value"><div class="profile-choice-group"><label class="profile-choice"><input type="radio" name="profileDisplayMode" value="real"> 真实姓名</label><label class="profile-choice"><input type="radio" name="profileDisplayMode" value="nickname"> 昵称</label><span class="profile-bind-note" id="profileDisplayModeText"></span><span id="profileNicknameOptionText" hidden></span></div></div></div>',
      '<div class="profile-row profile-row-top"><div class="profile-label">当前头像</div><div class="profile-value"><div class="profile-fieldline"><label class="profile-btn profile-file-btn">上传头像<input id="profileAvatarInput" type="file" accept="image/*"></label><button class="profile-btn secondary" id="profileResetAvatar" type="button">恢复默认头像</button></div><div class="profile-avatar-stage"><div class="profile-avatar-preview" id="profileAvatarPreview"></div></div><div class="profile-status" id="profileStatus" aria-live="polite"></div></div></div>',
      renderInfoRow('平台角色',escapeHtml(config.role)),
      renderInfoRow('人员编号',escapeHtml(config.personId||config.account)),
      renderInfoRow(config.organizationLabel||'系级',escapeHtml(config.organization||'未设置')),
      '</div></div>',
      '<div class="profile-block"><div class="profile-block-title">账号绑定</div><div class="profile-rows">',
      renderInfoRow('Email',escapeHtml(config.email||'None')),
      '<div class="profile-row"><div class="profile-label">手机号码</div><div class="profile-value">'+renderPhoneBinding()+'</div></div>',
      renderInfoRow('登录账号',escapeHtml(config.account)),
      '</div></div>',
      '<div class="profile-block"><div class="profile-block-title">'+escapeHtml(config.statsTitle)+'</div><div class="profile-rows"><div class="profile-data-grid">'+renderStats()+'</div><div class="profile-actions-row">'+renderActions()+'</div></div></div>',
      '</div>',
      '</div>'
    ].join('');
    bindProfileEvents();
    syncForm();
  }

  function renderInfoRow(label,value){
    return '<div class="profile-row"><div class="profile-label">'+escapeHtml(label)+'</div><div class="profile-value">'+value+'</div></div>';
  }

  function renderPhoneBinding(){
    if(config.phone)return '<span>'+escapeHtml(config.phone)+'</span>';
    return '<button class="profile-btn" type="button" data-profile-bind="phone">绑定</button>';
  }

  function renderStats(){
    if(!config.stats||!config.stats.length)return '<div class="profile-data-item"><span>暂无数据</span><b>-</b></div>';
    return config.stats.map(item=>'<div class="profile-data-item"><span>'+escapeHtml(item.label)+'</span><b>'+escapeHtml(item.value)+'</b></div>').join('');
  }

  function renderActions(){
    const actions=(config.actions||[]).map((action,index)=>{
      const cls='profile-btn'+(action.type==='primary'?'':action.type==='danger'?' danger':' secondary');
      return '<button class="'+cls+'" type="button" data-profile-action="'+index+'">'+escapeHtml(action.label)+'</button>';
    }).join('');
    const logout=config.allowLogout?'<button class="profile-btn danger" type="button" id="profileLogout">退出登录</button>':'';
    return actions+logout;
  }

  function bindProfileEvents(){
    const fileInput=document.getElementById('profileAvatarInput');
    const resetBtn=document.getElementById('profileResetAvatar');
    const saveBtn=document.getElementById('profileSave');
    const nickname=document.getElementById('profileNickname');
    if(fileInput)fileInput.addEventListener('change',handleAvatarUpload);
    if(resetBtn)resetBtn.addEventListener('click',resetAvatar);
    if(saveBtn)saveBtn.addEventListener('click',saveProfile);
    if(nickname)nickname.addEventListener('input',updateNicknameOptionText);
    document.querySelectorAll('input[name="profileDisplayMode"]').forEach(input=>{
      input.addEventListener('change',function(){
        const nickname=document.getElementById('profileNickname');
        state.nickname=(nickname?nickname.value:'').trim();
        state.displayMode=this.value;
        saveState();
        syncForm();
        applyIdentity();
      });
    });
    document.querySelectorAll('[data-profile-action]').forEach(button=>{
      button.addEventListener('click',function(){
        const action=(config.actions||[])[Number(this.dataset.profileAction)];
        runAction(action);
      });
    });
    document.querySelectorAll('[data-profile-bind]').forEach(button=>{
      button.addEventListener('click',function(){
        showStatus('手机号绑定已进入原型演示状态。','ok');
      });
    });
    const logout=document.getElementById('profileLogout');
    if(logout)logout.addEventListener('click',function(){
      if(typeof window.logout==='function')window.logout();
      else location.href='login1.1.html';
    });
  }

  function syncForm(){
    const nickname=document.getElementById('profileNickname');
    if(nickname)nickname.value=state.nickname||'';
    document.querySelectorAll('input[name="profileDisplayMode"]').forEach(input=>{
      input.checked=input.value===state.displayMode;
    });
    updateNicknameOptionText();
    updateProfilePreview();
  }

  function updateNicknameOptionText(){
    const option=document.getElementById('profileNicknameOptionText');
    const input=document.getElementById('profileNickname');
    const nickname=(input?input.value:state.nickname||'').trim();
    if(option)option.textContent=nickname||'未设置昵称时显示真实姓名';
  }

  function saveProfile(){
    const nicknameInput=document.getElementById('profileNickname');
    state.nickname=(nicknameInput?nicknameInput.value:'').trim();
    const checked=document.querySelector('input[name="profileDisplayMode"]:checked');
    state.displayMode=checked?checked.value:'real';
    if(state.nickname.length>20){
      showStatus('昵称不能超过 20 个字符。','bad');
      return;
    }
    saveState();
    syncForm();
    applyIdentity();
    showStatus('个人信息已保存。','ok');
  }

  function handleAvatarUpload(event){
    const file=event.target.files&&event.target.files[0];
    if(!file)return;
    if(!file.type||!file.type.startsWith('image/')){
      showStatus('请选择图片文件。','bad');
      event.target.value='';
      return;
    }
    if(file.size>2*1024*1024){
      showStatus('图片超过 2MB，请选择更小的头像。','bad');
      event.target.value='';
      return;
    }
    const reader=new FileReader();
    reader.onload=function(){
      state.avatarData=reader.result;
      if(saveState()){
        updateProfilePreview();
        applyIdentity();
        showStatus('头像已更新。','ok');
      }
    };
    reader.onerror=function(){showStatus('图片读取失败，请重新选择。','bad')};
    reader.readAsDataURL(file);
  }

  function resetAvatar(){
    state.avatarData='';
    saveState();
    updateProfilePreview();
    applyIdentity();
    showStatus('已恢复默认头像。','ok');
  }

  function updateProfilePreview(){
    const preview=document.getElementById('profileAvatarPreview');
    if(preview)setAvatar(preview);
    const displayName=document.getElementById('profileDisplayName');
    if(displayName)displayName.textContent=getDisplayName();
    const modeText=document.getElementById('profileDisplayModeText');
    if(modeText)modeText.textContent=state.displayMode==='nickname'?'当前以昵称展示':'当前以真实姓名展示';
  }

  function applyIdentity(){
    const displayName=getDisplayName();
    document.querySelectorAll('.user-name').forEach(el=>{el.textContent=displayName});
    document.querySelectorAll('.top .avatar').forEach(el=>setAvatar(el));
    document.querySelectorAll('.msg.user .msg-avatar').forEach(el=>setAvatar(el));
    const heroTitle=document.querySelector('#home .hero h1');
    if(config.updateGreeting&&heroTitle)heroTitle.textContent='欢迎回来，'+displayName;
    updateProfilePreview();
  }

  function setAvatar(el){
    if(state.avatarData){
      el.classList.add('has-image');
      el.style.backgroundImage='url("'+state.avatarData+'")';
      el.textContent='';
    }else{
      el.classList.remove('has-image');
      el.style.backgroundImage='';
      el.textContent=getAvatarText();
    }
  }

  function getDisplayName(){
    const nickname=(state.nickname||'').trim();
    return state.displayMode==='nickname'&&nickname?nickname:config.realName;
  }

  function getAvatarText(){
    const text=getDisplayName()||config.defaultAvatarText||'用';
    return Array.from(text.trim())[0]||config.defaultAvatarText||'用';
  }

  function showStatus(message,type){
    const status=document.getElementById('profileStatus');
    if(!status)return;
    status.className='profile-status '+(type||'');
    status.textContent=message;
  }

  function runAction(action){
    if(!action)return;
    if(action.target){
      if(typeof window.jump==='function')window.jump(action.target);
      else if(typeof window.show==='function')window.show(action.target,document.querySelector('.nav a[onclick*="'+action.target+'"]'));
      return;
    }
    showStatus(action.label+'已进入原型演示状态。','ok');
  }

  function escapeHtml(value){
    return String(value==null?'':value).replace(/[&<>"']/g,function(ch){
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch];
    });
  }

  function escapeAttr(value){
    return escapeHtml(value);
  }

  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);
  else init();
})();
