import streamlit as st
import pandas as pd
import hashlib, re
from datetime import datetime, date
from io import BytesIO
from supabase import create_client, Client

st.set_page_config(page_title='Company HR Contact Dashboard', page_icon='📞', layout='wide')

# ---------- Supabase ----------
@st.cache_resource
def get_supabase() -> Client:
    return create_client(st.secrets['SUPABASE_URL'], st.secrets['SUPABASE_KEY'])

sb = get_supabase()

def hp(p): return hashlib.sha256(p.encode('utf-8')).hexdigest()
def rows(resp): return resp.data or []
def df(resp): return pd.DataFrame(rows(resp))
def one(table, **eq):
    q=sb.table(table).select('*')
    for k,v in eq.items(): q=q.eq(k,v)
    r=rows(q.limit(1).execute())
    return r[0] if r else None

def users(active_only=True):
    q=sb.table('users').select('id,username,display_name,role,active').order('display_name')
    if active_only: q=q.eq('active',True)
    return pd.DataFrame(rows(q.execute()))

def company(cid): return one('companies', id=cid)
def authenticate(username, password):
    try:
        result = (
            sb.table("users")
            .select("id,username,display_name,password_hash,role,active")
            .ilike("username", username.strip())
            .eq("active", True)
            .limit(1)
            .execute()
        )

        if not result.data:
            return None

        user = result.data[0]

        if user["password_hash"] == hp(password):
            return user

        return None

    except Exception as e:
        st.error("SUPABASE ERROR:")
        st.code(str(e))
        return None

def kamaljit():
    r=rows(sb.table('users').select('id,username,display_name').ilike('username','kamaljit').eq('active',True).limit(1).execute())
    return r[0] if r else None

def notify(uid,title,message):
    sb.table('portal_notifications').insert({'recipient_user_id':uid,'title':title,'message':message}).execute()

def joined_responses(uid=None):
    q=sb.table('responses').select('id,company_id,user_id,response_status,remarks,call_date,followup_date,followup_action,created_at,companies(company,hr_name,phone,email),users(display_name,username)')
    if uid is not None: q=q.eq('user_id',uid)
    data=rows(q.order('created_at',desc=True).execute())
    out=[]
    for r in data:
        c=r.get('companies') or {}; u=r.get('users') or {}
        out.append({'id':r['id'],'Company':c.get('company'),'Current_HR':c.get('hr_name'),'Phone':c.get('phone'),'Called_By':u.get('display_name'),'Response':r.get('response_status'),'Remarks':r.get('remarks'),'Action_Taken':r.get('followup_action'),'Call_Date':r.get('call_date'),'Follow_Up':r.get('followup_date'),'Logged_At':r.get('created_at'),'user_id':r.get('user_id'),'company_id':r.get('company_id')})
    return pd.DataFrame(out)

def export_excel():
    cs=df(sb.table('companies').select('*').order('company').execute())
    rs=joined_responses()
    hs=df(sb.table('hr_history').select('*').order('changed_at',desc=True).execute())
    out=BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        cs.to_excel(w,'Companies',index=False); rs.to_excel(w,'Responses',index=False); hs.to_excel(w,'HR History',index=False)
    out.seek(0); return out.getvalue()

# ---------- Login ----------
if 'user' not in st.session_state: st.session_state.user=None
if st.session_state.user is None:
    st.title('📞 Company HR Contact Dashboard')
    with st.form('login'):
        un=st.text_input('Username'); pw=st.text_input('Password',type='password'); go=st.form_submit_button('Login',use_container_width=True)
    if go:
        u=authenticate(un,pw)
        if u: st.session_state.user=u; st.rerun()
        else: st.error('Invalid username or password.')
    st.stop()

user=st.session_state.user; master=user['role']=='master'; uid=user['id']
st.sidebar.title('📞 HR Calling Portal'); st.sidebar.write(f"Logged in as **{user['display_name']}**")
pages=['Dashboard','My Companies','Follow-ups','Call History','Share with Kamaljit','Change Password','Add Contacts from Excel']
if master: pages=['Dashboard','Company Database','Assignments','Follow-ups','Call History','Reports','User Management','Share with Kamaljit','Change Password','Add Contacts from Excel']
page=st.sidebar.radio('Navigate',pages)
notes=df(sb.table('portal_notifications').select('*').eq('recipient_user_id',uid).is_('read_at','null').order('created_at',desc=True).execute())
if not notes.empty:
    st.sidebar.divider(); st.sidebar.subheader(f'🔔 Portal Messages ({len(notes)})')
    for _,n in notes.iterrows(): st.sidebar.warning(f"**{n['title']}**\n\n{n['message']}")
    if st.sidebar.button('✓ Mark messages read'):
        sb.table('portal_notifications').update({'read_at':datetime.now().isoformat()}).eq('recipient_user_id',uid).is_('read_at','null').execute(); st.rerun()
if st.sidebar.button('Logout',use_container_width=True): st.session_state.user=None; st.rerun()

# ---------- Dashboard ----------
if page=='Dashboard':
    st.title('Dashboard')
    cs=df(sb.table('companies').select('id,assigned_to').eq('active',True).execute()); rs=df(sb.table('responses').select('id,company_id,user_id,followup_date').execute())
    if master:
        a=int(cs['assigned_to'].notna().sum()) if not cs.empty else 0
        c1,c2,c3,c4=st.columns(4); c1.metric('Active Companies',len(cs)); c2.metric('Assigned',a); c3.metric('Total Responses',len(rs)); c4.metric('Follow-ups Today',int((rs.get('followup_date',pd.Series(dtype=str))==date.today().isoformat()).sum()))
    else:
        mine=cs[cs.assigned_to==uid] if not cs.empty else cs; myr=rs[rs.user_id==uid] if not rs.empty else rs
        called=myr.company_id.nunique() if not myr.empty else 0; today=date.today().isoformat()
        c1,c2,c3,c4=st.columns(4); c1.metric('My Assigned Companies',len(mine)); c2.metric('Companies Contacted',called); c3.metric('Follow-ups Today',int((myr.get('followup_date',pd.Series(dtype=str))==today).sum())); c4.metric('Remaining',max(len(mine)-called,0))

elif page=='Company Database' and master:
    st.title('🏢 Company Database')
    data=df(sb.table('companies').select('*').eq('active',True).order('company').execute()); search=st.text_input('Search company / sector / HR / email')
    if search and not data.empty:
        mask=data.fillna('').astype(str).apply(lambda x:x.str.contains(search,case=False,regex=False)).any(axis=1); data=data[mask]
    st.dataframe(data,use_container_width=True,hide_index=True)
    if not data.empty:
        opts={f"{r['company']} — {r.get('hr_name') or 'No HR'}":int(r['id']) for _,r in data.iterrows()}; lab=st.selectbox('Select company to edit',opts); c=company(opts[lab])
        with st.form('editc'):
            cn=st.text_input('Company',c.get('company') or ''); sec=st.text_input('Sector',c.get('sector') or ''); hr=st.text_input('HR',c.get('hr_name') or ''); ph=st.text_input('Phone',c.get('phone') or ''); em=st.text_input('Email',c.get('email') or ''); save=st.form_submit_button('Save Changes')
        if save:
            if any([(c.get('hr_name') or '')!=hr,(c.get('phone') or '')!=ph,(c.get('email') or '')!=em]): sb.table('hr_history').insert({'company_id':c['id'],'old_hr_name':c.get('hr_name'),'old_phone':c.get('phone'),'old_email':c.get('email'),'changed_by':uid}).execute()
            sb.table('companies').update({'company':cn,'sector':sec,'hr_name':hr,'phone':ph,'email':em}).eq('id',c['id']).execute(); st.success('Updated.'); st.rerun()
    st.subheader('Add New Company')
    with st.form('addc'):
        cn=st.text_input('Company name'); sec=st.text_input('Sector'); hr=st.text_input('HR name'); ph=st.text_input('Phone'); em=st.text_input('Email'); add=st.form_submit_button('Add Company')
    if add:
        if not cn.strip(): st.error('Company name is required.')
        else: sb.table('companies').insert({'company':cn.strip(),'sector':sec,'hr_name':hr,'phone':ph,'email':em}).execute(); st.success('Company added.'); st.rerun()

elif page=='Assignments' and master:
    st.title('👑 Assign Companies'); us=users(); us=us[us.role=='admin']; amap={f"{r.display_name} ({r.username})":int(r.id) for _,r in us.iterrows()}
    cs=df(sb.table('companies').select('id,company,sector,hr_name,assigned_to').eq('active',True).order('company').execute()); mode=st.radio('Show',['Unassigned','All'],horizontal=True)
    if mode=='Unassigned' and not cs.empty: cs=cs[cs.assigned_to.isna()]
    labels={f"{r.company} | {r.sector or ''} | HR: {r.hr_name or ''}":int(r.id) for _,r in cs.iterrows()}; sel=st.multiselect('Select companies',labels); an=st.selectbox('Assign to',list(amap))
    if st.button('Assign Selected',type='primary'):
        for x in sel: sb.table('companies').update({'assigned_to':amap[an]}).eq('id',labels[x]).execute()
        st.success(f'Assigned {len(sel)} company/companies.'); st.rerun()

elif page=='My Companies' and not master:
    st.title('📞 My Assigned Companies'); data=df(sb.table('companies').select('*').eq('active',True).eq('assigned_to',uid).order('company').execute()); st.dataframe(data,use_container_width=True,hide_index=True)
    if not data.empty:
        opts={f"{r.company} — {r.hr_name or 'No HR'}":int(r.id) for _,r in data.iterrows()}; lab=st.selectbox('Select company to record call',opts); c=company(opts[lab]); st.write(f"**Phone:** {c.get('phone') or '-'}   **Email:** {c.get('email') or '-'}")
        # similar company warning across other coordinators
        hist=joined_responses(); selected=str(c.get('company') or '').casefold(); toks={t for t in re.findall(r'[a-z0-9]+',selected) if len(t)>=4}
        if not hist.empty:
            other=hist[hist.user_id!=uid]
            for _,r in other.iterrows():
                pn=str(r.Company or '').casefold(); pt={t for t in re.findall(r'[a-z0-9]+',pn) if len(t)>=4}
                if selected in pn or pn in selected or toks & pt:
                    st.warning(f"⚠️ Similar company already contacted by **{r.Called_By}**. Company: {r.Company} | Status: {r.Response} | Remarks: {r.Remarks}"); break
        statuses=['Interested','Not Interested','Call Later','No Response','Wrong Number','HR Changed','Email Sent','Meeting Scheduled','Follow-up Required','Other']
        with st.form('response'):
            status=st.selectbox('Response',statuses); remarks=st.text_area('Remarks *'); cd=st.date_input('Call date',date.today()); nf=st.checkbox('Set follow-up reminder'); fd=st.date_input('Next follow-up date',date.today()) if nf else None; save=st.form_submit_button('Save Response',type='primary')
        if save:
            if not remarks.strip(): st.error('Remarks are compulsory.')
            else: sb.table('responses').insert({'company_id':c['id'],'user_id':uid,'response_status':status,'remarks':remarks.strip(),'call_date':cd.isoformat(),'followup_date':fd.isoformat() if fd else None}).execute(); st.success('Response saved.'); st.rerun()

elif page=='Follow-ups':
    st.title('⏰ Follow-ups'); h=joined_responses(None if master else uid)
    if h.empty: st.info('No follow-ups scheduled.')
    else:
        h=h[h.Follow_Up.notna()].copy(); h['Follow_Up_Date']=pd.to_datetime(h.Follow_Up,errors='coerce').dt.date; today=date.today(); h['Status']=h.Follow_Up_Date.apply(lambda d:'🔴 Overdue' if d<today else ('🟠 Due Today' if d==today else '🟡 Upcoming')); st.dataframe(h,use_container_width=True,hide_index=True)
        due=h[h.Follow_Up_Date<=today]
        if not due.empty:
            st.subheader('📌 Action Taken for Due Follow-ups'); actions=['Call completed','Information shared','Meeting scheduled','Interested - further discussion','Not interested','No response - call again','HR/contact changed','Follow-up rescheduled','Follow-up closed','Other']
            for _,r in due.iterrows():
                a=st.selectbox(f"{r.Company} — {r.Follow_Up}",['-- Select action --']+actions,key=f'a{r.id}')
                if st.button('Save Action Taken',key=f's{r.id}'):
                    if a.startswith('--'): st.error('Select an action.')
                    else: sb.table('responses').update({'followup_action':a}).eq('id',int(r.id)).execute(); st.success('Saved.'); st.rerun()

elif page=='Call History':
    st.title('📝 Call Response History'); h=joined_responses()
    if not master and not h.empty: h=h[h.user_id==uid]
    if master and not h.empty:
        names=['All']+sorted(h.Called_By.dropna().unique().tolist()); who=st.selectbox('Faculty Coordinator',names)
        if who!='All': h=h[h.Called_By==who]
    st.dataframe(h,use_container_width=True,hide_index=True,height=520)
    if not h.empty:
        out=BytesIO(); h.to_excel(out,index=False,engine='openpyxl'); st.download_button('📥 Download Call History',out.getvalue(),'Call_History.xlsx')

elif page=='Reports' and master:
    st.title('📊 Reports'); h=joined_responses()
    if not h.empty:
        st.subheader('Response Distribution'); st.bar_chart(h.Response.value_counts())
        perf=h.groupby('Called_By').agg(Calls=('id','count'),Companies=('company_id','nunique')).reset_index(); st.dataframe(perf,use_container_width=True,hide_index=True)
    st.download_button('⬇️ Download Full Excel Report',export_excel(),f"company_calling_report_{date.today().isoformat()}.xlsx")

elif page=='User Management' and master:
    st.title('👥 User Management'); us=users(False); st.dataframe(us,use_container_width=True,hide_index=True)
    st.subheader('➕ Add New Faculty Coordinator')
    with st.form('addu'):
        un=st.text_input('Username'); dn=st.text_input('Faculty Coordinator Name'); pw=st.text_input('Temporary Password',type='password'); add=st.form_submit_button('Create User')
    if add:
        if not un.strip() or not dn.strip(): st.error('Username and name are required.')
        elif len(pw)<8: st.error('Password must contain at least 8 characters.')
        else:
            try: sb.table('users').insert({'username':un.strip().lower(),'display_name':dn.strip(),'password_hash':hp(pw),'role':'admin','active':True}).execute(); st.success('User created.'); st.rerun()
            except Exception as e: st.error(f'Could not create user. Username may already exist. {e}')
    st.subheader('🔐 Change Any User Username / Password')
    labels={f"{r.display_name} ({r.username})":int(r.id) for _,r in us.iterrows()}; lab=st.selectbox('Select user',labels); target=us[us.id==labels[lab]].iloc[0]
    with st.form('editcred'):
        nun=st.text_input('Username',value=target.username); npw=st.text_input('New password (leave blank to keep current)',type='password'); cpw=st.text_input('Confirm new password',type='password'); save=st.form_submit_button('Save Credentials')
    if save:
        if not nun.strip(): st.error('Username cannot be blank.')
        elif npw and len(npw)<8: st.error('Password must be at least 8 characters.')
        elif npw!=cpw: st.error('Passwords do not match.')
        else:
            payload={'username':nun.strip().lower()}
            if npw: payload['password_hash']=hp(npw)
            try: sb.table('users').update(payload).eq('id',int(target.id)).execute(); st.success('Credentials updated.'); st.rerun()
            except Exception as e: st.error(f'Could not update credentials: {e}')

elif page=='Share with Kamaljit':
    st.title('📨 Share HR Contact with Kamaljit'); cs=df(sb.table('companies').select('id,company,hr_name,email').eq('active',True).order('company').execute())
    if not master and not cs.empty:
        assigned=df(sb.table('companies').select('id').eq('assigned_to',uid).eq('active',True).execute()); ids=set(assigned.id.tolist()) if not assigned.empty else set(); cs=cs[cs.id.isin(ids)]
    if cs.empty: st.info('No companies available.')
    else:
        opts={f"{r.company} — {r.hr_name or 'No HR'}":int(r.id) for _,r in cs.iterrows()}; lab=st.selectbox('Company',opts); c=company(opts[lab])
        with st.form('share'):
            email=st.text_input('HR Email ID',value=c.get('email') or ''); remarks=st.text_area('Remarks (optional)'); share=st.form_submit_button('Share with Kamaljit',type='primary')
        if share:
            k=kamaljit()
            if not email.strip(): st.error('HR Email ID is required.')
            elif not k: st.error("Active username 'kamaljit' was not found.")
            else:
                sb.table('contact_shares').insert({'company_id':c['id'],'company_name':c['company'],'hr_email':email.strip(),'shared_by':uid,'shared_with':k['id'],'remarks':remarks.strip() or None}).execute()
                notify(k['id'],'📨 New HR Contact Shared',f"Company: {c['company']}\nHR Email: {email.strip()}\nShared by: {user['display_name']}\nDate: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}")
                st.success('Company and HR email shared with Kamaljit. No email was sent to the HR.')
    if user['username'].casefold()=='kamaljit':
        st.subheader('Contacts Shared With Me'); shares=df(sb.table('contact_shares').select('*').eq('shared_with',uid).order('shared_at',desc=True).execute()); st.dataframe(shares,use_container_width=True,hide_index=True)

elif page=='Change Password':
    st.title('🔐 Change Password')
    with st.form('cp'):
        old=st.text_input('Current password',type='password'); new=st.text_input('New password',type='password'); con=st.text_input('Confirm new password',type='password'); change=st.form_submit_button('Change Password')
    if change:
        if authenticate(user['username'],old) is None: st.error('Current password is incorrect.')
        elif len(new)<8: st.error('New password must be at least 8 characters.')
        elif new!=con: st.error('Passwords do not match.')
        else: sb.table('users').update({'password_hash':hp(new)}).eq('id',uid).execute(); st.success('Password changed.')

elif page=='Add Contacts from Excel':
    st.title('➕ Add Contacts from Excel'); f=st.file_uploader('Choose Excel file',type=['xlsx','xls'])
    if f:
        nd=pd.read_excel(f); required=['Company','Sector','HR','Phno','Email']; missing=[x for x in required if x not in nd.columns]
        if missing: st.error('Missing columns: '+', '.join(missing))
        else:
            nd=nd[required].copy(); nd['Company']=nd.Company.fillna('').astype(str).str.strip(); nd=nd[nd.Company!='']; st.dataframe(nd.head(20),use_container_width=True)
            if st.button('➕ Merge Contacts',type='primary'):
                added=existing=0
                current=df(sb.table('companies').select('*').execute()); cmap={str(r.company).strip().casefold():r for _,r in current.iterrows()} if not current.empty else {}
                for _,r in nd.iterrows():
                    name=str(r.Company).strip(); key=name.casefold(); vals={'sector':'' if pd.isna(r.Sector) else str(r.Sector).strip(),'hr_name':'' if pd.isna(r.HR) else str(r.HR).strip(),'phone':'' if pd.isna(r.Phno) else str(r.Phno).strip(),'email':'' if pd.isna(r.Email) else str(r.Email).strip()}
                    if key in cmap:
                        old=cmap[key]; payload={k:v for k,v in vals.items() if v and not str(old.get(k) or '').strip()}
                        if payload: sb.table('companies').update(payload).eq('id',int(old.id)).execute()
                        existing+=1
                    else: sb.table('companies').insert({'company':name,**vals,'active':True}).execute(); added+=1
                st.success(f'Merged successfully. New: {added}; Existing: {existing}'); st.rerun()

# completion summary for coordinators
if not master:
    st.divider(); assigned=df(sb.table('companies').select('id').eq('assigned_to',uid).eq('active',True).execute()); rr=df(sb.table('responses').select('company_id').eq('user_id',uid).execute()); total=len(assigned); called=rr.company_id.nunique() if not rr.empty else 0
    c1,c2,c3=st.columns(3); c1.metric('Assigned Companies',total); c2.metric('Companies Called',called); c3.metric('Remaining',max(total-called,0))
