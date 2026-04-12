import random
import csv
from django.http import HttpResponse
from datetime import date, timedelta
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db.models import Q
from .models import Profile, MonthlyMeet, Match, Notice
from django.contrib.auth.models import User
from django.contrib.auth import login, authenticate, logout
from django.contrib.auth import update_session_auth_hash # 비밀번호 변경 후 로그아웃 방지
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font


# --- (신규) 권한 체크용 헬퍼 함수 ---
def is_manager(user):
    """총관리자(is_superuser)이거나 사장님(is_owner)인지 확인"""
    if not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    if hasattr(user, 'profile') and user.profile.is_owner:
        return True
    return False

# 1. 핸디캡 계산 엔진 (기존 유지 - 기본 2점)
def calculate_handicap_logic(p1, p2):
    tier_weight = {'A': 3, 'B': 2, 'C': 1}
    
    if tier_weight[p1.group] < tier_weight[p2.group]:
        p1, p2 = p2, p1
        
    diff = tier_weight[p1.group] - tier_weight[p2.group]
    
    base_handicap = 0
    if diff == 1: base_handicap = 2   # A vs B, B vs C
    elif diff == 2: base_handicap = 6 # A vs C

    gender_bonus = 0
    if p1.gender == 'M' and p2.gender == 'F':
        gender_bonus = 2
    elif p1.gender == 'F' and p2.gender == 'M':
        gender_bonus = -2

    final_handicap = max(0, base_handicap + gender_bonus)
    
    last_month_limit = date.today() - timedelta(days=40)
    rematch_exists = Match.objects.filter(
        meet__date__gte=last_month_limit,
        is_completed=True
    ).filter(
        Q(player1=p1, player2=p2) | Q(player1=p2, player2=p1)
    ).exists()

    if rematch_exists:
        final_handicap = max(0, final_handicap - 1)
    
    return final_handicap

# 2. 그룹별 승률 계산 헬퍼 함수
def get_top_players(group_name):
    profiles = Profile.objects.filter(group=group_name)
    stats = []
    for p in profiles:
        # 'meet__is_finalized=True' 조건을 추가하여 마감된 경기만 집계
        matches = Match.objects.filter(
            Q(player1=p) | Q(player2=p), 
            is_completed=True,
            meet__is_finalized=True 
        )
        total = matches.count()
        wins = 0
        for m in matches:
            if m.player1 == p and (m.p1_score or 0) > (m.p2_score or 0): wins += 1
            if m.player2 == p and (m.p2_score or 0) > (m.p1_score or 0): wins += 1
            
        win_rate = int((wins / total * 100)) if total > 0 else 0
        if total > 0:
            stats.append({'profile': p, 'win_rate': win_rate, 'total': total})
            
    stats.sort(key=lambda x: x['win_rate'], reverse=True)
    return stats[:3]

# 3. 메인 포털 대시보드 뷰
@login_required(login_url='login')
def dashboard(request):
    meet = MonthlyMeet.objects.order_by('-date').first()
    all_profiles = Profile.objects.all()
    
    notices = Notice.objects.all().order_by('-is_important', '-created_at')[:5]
    top_a = get_top_players('A')
    top_b = get_top_players('B')
    top_c = get_top_players('C')

    # 전체 인원의 상세 랭킹 데이터 계산
    all_stats_sorted = []
    for p in all_profiles:
        p_matches = Match.objects.filter(Q(player1=p) | Q(player2=p), is_completed=True)
        p_total = p_matches.count()
        
        p_wins = sum(1 for m in p_matches if (m.player1 == p and (m.p1_score or 0) > (m.p2_score or 0)) or (m.player2 == p and (m.p2_score or 0) > (m.p1_score or 0)))
        p_losses = p_total - p_wins
        p_win_rate = int((p_wins / p_total * 100)) if p_total > 0 else 0
        
        all_stats_sorted.append({
            'profile': p,
            'win_rate': p_win_rate,
            'wins': p_wins,
            'losses': p_losses
        })
        
    # 승률을 기준으로 내림차순 정렬
    all_stats_sorted.sort(key=lambda x: x['win_rate'], reverse=True)
    
    # 🌟 추가된 부분: 조별 랭킹 탭에서 순위를 1번부터 매기기 위해 리스트를 분리
    all_a = [s for s in all_stats_sorted if s['profile'].group == 'A']
    all_b = [s for s in all_stats_sorted if s['profile'].group == 'B']
    all_c = [s for s in all_stats_sorted if s['profile'].group == 'C']

    user_stat = None
    if request.user.is_authenticated and hasattr(request.user, 'profile'):
        p = request.user.profile
        matches = Match.objects.filter(Q(player1=p) | Q(player2=p), is_completed=True)
        total = matches.count()
        wins = sum(1 for m in matches if (m.player1 == p and (m.p1_score or 0) > (m.p2_score or 0)) or (m.player2 == p and (m.p2_score or 0) > (m.p1_score or 0)))
        win_rate = int((wins / total * 100)) if total > 0 else 0
        user_stat = {'win_rate': win_rate, 'total': total, 'wins': wins}

    login_failed_username = request.session.pop('login_failed_username', '')

    context = {
        'meet': meet,
        'all_profiles': all_profiles,
        'all_stats_sorted': all_stats_sorted, 
        'all_a': all_a, # 🌟 분리된 리스트 HTML로 전달
        'all_b': all_b, 
        'all_c': all_c, 
        'notices': notices,
        'top_a': top_a, 
        'top_b': top_b, 
        'top_c': top_c,
        'user_stat': user_stat,
        'is_manager': is_manager(request.user),
        'login_failed_username': login_failed_username,
    }

    if meet:
        context['matches_2'] = Match.objects.filter(meet=meet, court=2, is_completed=False)
        context['matches_3'] = Match.objects.filter(meet=meet, court=3, is_completed=False)
        context['completed'] = Match.objects.filter(meet=meet, is_completed=True)
        context['all_matches'] = Match.objects.filter(meet=meet)

    return render(request, 'matches/dashboard.html', context)

# 4. 팝업창 연동: 모임 생성 및 랜덤 배치 (코트 선택 기능 추가)
@login_required
def create_meet_and_matches(request):
    if request.method == "POST" and is_manager(request.user):
        title = request.POST.get('title')
        meet_date = request.POST.get('date')
        selected_profile_ids = request.POST.getlist('profiles') 
        
        # 신규: HTML에서 체크된 코트 정보 가져오기 (예: ['2', '3'] 또는 ['2'])
        selected_courts = request.POST.getlist('courts') 

        # 방어 로직: 코트를 하나도 선택 안 했을 경우
        if not selected_courts:
            messages.error(request, "최소 하나의 코트를 선택해야 합니다.")
            return redirect('dashboard')

        # 문자열 리스트를 숫자 리스트로 변환
        available_courts = [int(c) for c in selected_courts]

        if title and meet_date and selected_profile_ids:
            meet = MonthlyMeet.objects.create(title=title, date=meet_date)

            players = list(Profile.objects.filter(id__in=selected_profile_ids))
            random.shuffle(players)

            for i in range(0, len(players) - 1, 2):
                p1, p2 = players[i], players[i+1]
                handicap = calculate_handicap_logic(p1, p2)
                
                Match.objects.create(
                    meet=meet,
                    player1=p1,
                    player2=p2,
                    applied_handicap=handicap,
                    # 신규: 선택된 코트 안에서만 랜덤 배정되도록 수정!
                    court=random.choice(available_courts) 
                )
            messages.success(request, f"'{title}' 모임과 대진표가 성공적으로 생성되었습니다!")
            
    return redirect('dashboard')

# 5. 점수 입력 (일반유저 1회 제한, 총관리자/사장님 무제한 수정 권한 적용)
@login_required
def record_score(request, match_id):
    match = get_object_or_404(Match, id=match_id)
    
    # 🌟 수정: 일반 유저가 완료된 경기 수정을 시도할 경우 차단 및 안내
    if match.is_completed and not is_manager(request.user):
        messages.error(request, "점수 수정은 사장님께 문의하세요.")
        return redirect('dashboard')
        
    if request.method == 'POST':
        match.p1_score = request.POST.get('p1_score')
        match.p2_score = request.POST.get('p2_score')
        match.is_completed = True
        
        # 🌟 추가: 최초 점수 입력자 저장 (이미 완료된 경기를 사장님이 수정할 땐 최초 입력자 유지)
        if not match.recorded_by:
            match.recorded_by = request.user
            
        match.save()
        messages.success(request, "점수가 성공적으로 기록되었습니다.")
        return redirect('dashboard')
    
# 6. 회원가입 함수
def signup(request):
    # 이미 로그인한 사람이 회원가입 페이지로 오면 메인으로 돌려보냄
    if request.user.is_authenticated:
        return redirect('dashboard')

    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        name = request.POST.get('name')
        group = request.POST.get('group')
        gender = request.POST.get('gender')

        # 아이디 중복 체크 방어 로직
        if User.objects.filter(username=username).exists():
            messages.error(request, '이미 사용 중인 아이디입니다.')
            return redirect('signup')

        # 1. Django 기본 User 계정 생성
        user = User.objects.create_user(username=username, password=password)
        
        # 2. 동호회 프로필(Profile) 생성 및 연결 (일반 가입자는 is_owner가 디폴트인 False로 들어감)
        Profile.objects.create(user=user, name=name, group=group, gender=gender)

        # 3. 가입 완료와 동시에 자동 로그인 처리 후 대시보드로 이동
        login(request, user)
        return redirect('dashboard')

    return render(request, 'matches/signup.html')

# 7. 회원 즉시 추가 기능 (총관리자/사장님 권한 적용)
@login_required
def add_member_by_admin(request):
    if request.method == 'POST' and is_manager(request.user):
        username = request.POST.get('username')
        password = request.POST.get('password')
        name = request.POST.get('name')
        group = request.POST.get('group')
        gender = request.POST.get('gender')

        if User.objects.filter(username=username).exists():
            messages.error(request, f"'{username}'(은)는 이미 사용 중인 아이디입니다.")
        else:
            # 1. 계정 생성
            user = User.objects.create_user(username=username, password=password)
            # 2. 프로필 생성
            Profile.objects.create(user=user, name=name, group=group, gender=gender)
            messages.success(request, f"🎉 {name} 회원({group}조)이 성공적으로 추가되었습니다!")
            
    return redirect('dashboard')

# (신규) 9. 오늘 모임 최종 마감 함수
@login_required
def finalize_meet(request, meet_id):
    if is_manager(request.user):
        meet = get_object_or_404(MonthlyMeet, id=meet_id)
        meet.is_finalized = True
        meet.save()
        messages.success(request, f"🎉 {meet.title} 일정이 마감되었습니다. 랭킹에 반영됩니다.")
    return redirect('dashboard')

# 10. 엑셀(CSV) 템플릿 다운로드 함수
@login_required
def download_member_template(request):
    if is_manager(request.user):
        # 엑셀에서 한글이 깨지지 않도록 utf-8-sig 포맷 사용
        response = HttpResponse(content_type='text/csv; charset=utf-8-sig')
        response['Content-Disposition'] = 'attachment; filename="member_upload_template.csv"'
        
        writer = csv.writer(response)
        # 1행: 헤더 (안내문)
        writer.writerow(['접속아이디', '초기비밀번호', '실명', '조(A/B/C)', '성별(M/F)'])
        # 2행: 예시 데이터
        writer.writerow(['user01', '1234', '홍길동', 'B', 'M'])
        writer.writerow(['user02', '1234', '김유신', 'A', 'M'])
        writer.writerow(['user03', '1234', '유관순', 'C', 'F'])
        
        return response
    return redirect('dashboard')

# 11. 엑셀(CSV) 일괄 업로드 처리 함수
@login_required
def upload_members_bulk(request):
    if request.method == 'POST' and is_manager(request.user):
        csv_file = request.FILES.get('excel_file')
        
        if not csv_file or not csv_file.name.endswith('.csv'):
            messages.error(request, "CSV 파일을 업로드해주세요.")
            return redirect('dashboard')

        try:
            # 파일 읽기 및 파싱
            decoded_file = csv_file.read().decode('utf-8-sig').splitlines()
            reader = csv.reader(decoded_file)
            next(reader) # 첫 번째 줄(헤더)은 건너뜀
            
            success_count = 0
            for row in reader:
                if len(row) >= 5:
                    # 엑셀 데이터 추출 후 앞뒤 공백 제거
                    username = row[0].strip()
                    password = row[1].strip()
                    name = row[2].strip()
                    group = row[3].strip().upper()
                    gender = row[4].strip().upper()
                    
                    # 아이디가 존재하지 않을 때만 생성 (중복 방지)
                    if username and not User.objects.filter(username=username).exists():
                        user = User.objects.create_user(username=username, password=password)
                        Profile.objects.create(user=user, name=name, group=group, gender=gender)
                        success_count += 1
                        
            messages.success(request, f"🎉 총 {success_count}명의 회원이 성공적으로 일괄 추가되었습니다!")
        except Exception as e:
            messages.error(request, f"파일 처리 중 오류가 발생했습니다. 양식을 확인해주세요. (에러: {e})")
            
    return redirect('dashboard')

# 12. 경기 개별 수정 (선수, 코트 변경 등 - 사장님 전용)
@login_required
def update_match_detail(request, match_id):
    if request.method == "POST" and is_manager(request.user):
        match = get_object_or_404(Match, id=match_id)
        
        p1_id = request.POST.get('player1')
        p2_id = request.POST.get('player2')
        court = request.POST.get('court')
        
        # 선수 변경 반영
        match.player1 = get_object_or_404(Profile, id=p1_id)
        match.player2 = get_object_or_404(Profile, id=p2_id)
        match.court = int(court)
        
        # 점수가 함께 넘어온 경우 (수정 모드)
        p1_score = request.POST.get('p1_score')
        p2_score = request.POST.get('p2_score')
        if p1_score is not None and p2_score is not None:
            match.p1_score = int(p1_score)
            match.p2_score = int(p2_score)
            match.is_completed = True
            
        # 핸디캡 재계산 (선수가 바뀌었을 수 있으므로)
        match.applied_handicap = calculate_handicap_logic(match.player1, match.player2)
        match.save()
        
        messages.success(request, "경기 정보가 수정되었습니다.")
    return redirect('dashboard')

# 13. 독립된 전용 로그인 화면 뷰 (기존 custom_login 덮어쓰기)
def custom_login(request):
    if request.user.is_authenticated:
        return redirect('dashboard')

    if request.method == 'POST':
        user_id = request.POST.get('username')
        user_pw = request.POST.get('password')
        
        user = authenticate(request, username=user_id, password=user_pw)
        
        if user is not None:
            login(request, user)
            if 'login_failed_username' in request.session:
                del request.session['login_failed_username']
            return redirect('dashboard')
        else:
            request.session['login_failed_username'] = user_id
            messages.error(request, "아이디 또는 패스워드가 일치하지 않습니다.", extra_tags="login_error")
            return redirect('login') 
            
    login_failed_username = request.session.pop('login_failed_username', '')
    return render(request, 'matches/login.html', {'login_failed_username': login_failed_username})


# 14. 커스텀 로그아웃 (신규 추가)
def custom_logout(request):
    logout(request)
    return redirect('login')

# 17. 비밀번호 변경 기능
@login_required
def change_password(request):
    if request.method == 'POST':
        old_pw = request.POST.get('old_password')
        new_pw = request.POST.get('new_password')
        user = request.user
        
        if user.check_password(old_pw):
            user.set_password(new_pw)
            user.save()
            # 🌟 비밀번호가 바뀌어도 현재 세션을 유지(자동 로그아웃 방지)
            update_session_auth_hash(request, user)
            messages.success(request, "비밀번호가 성공적으로 변경되었습니다.")
        else:
            messages.error(request, "현재 비밀번호가 일치하지 않습니다.")
            
    return redirect('dashboard')

# 18. 경기 일정 엑셀 다운로드 (세로형) - 사장님 전용
@login_required
def export_schedule_vertical(request, meet_id):
    if not is_manager(request.user):
        return redirect('dashboard')

    meet = get_object_or_404(MonthlyMeet, id=meet_id)
    wb = Workbook()
    ws = wb.active
    ws.title = "세로형 일정"

    # 스타일 설정
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True, size=12)
    court_font = Font(bold=True, size=14, color="0000FF") # 코트명 강조
    center_align = Alignment(horizontal='center', vertical='center')

    current_row = 1
    # 처리할 코트 번호들
    for court_num in [2, 3]:
        matches = Match.objects.filter(meet=meet, court=court_num).order_by('id')
        if not matches.exists(): continue # 경기가 없는 코트는 건너뜀

        # 코트 제목 생성
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws.cell(row=current_row, column=1, value=f"[{court_num} 코트 일정]")
        cell.font = court_font
        cell.alignment = center_align
        current_row += 1

        # 헤더 생성
        headers = ['순서', '선수 1', 'vs', '선수 2', '비고']
        for col, val in enumerate(headers, 1):
            c = ws.cell(row=current_row, column=col, value=val)
            c.font = header_font
            c.border = thin_border
            c.alignment = center_align
        current_row += 1

        # 데이터 입력 및 핸디캡 계산
        for idx, m in enumerate(matches, 1):
            p1_display = m.player1.name
            p2_display = m.player2.name
            
            # 핸디캡 수혜자 찾기 로직 (티어 차이 우선, 성별 차이 후순위)
            tier_weight = {'A': 3, 'B': 2, 'C': 1}
            h_val = m.applied_handicap
            if h_val > 0:
                if tier_weight[m.player1.group] < tier_weight[m.player2.group]:
                    p1_display += f" (+{h_val})"
                elif tier_weight[m.player1.group] > tier_weight[m.player2.group]:
                    p2_display += f" (+{h_val})"
                else: # 티어가 같으면 여성 선수가 보너스를 받는 구조로 표시
                    if m.player1.gender == 'F' and m.player2.gender == 'M':
                        p1_display += f" (+{h_val})"
                    else:
                        p2_display += f" (+{h_val})"

            row_data = [idx, p1_display, 'vs', p2_display, '']
            for col, val in enumerate(row_data, 1):
                c = ws.cell(row=current_row, column=col, value=val)
                c.border = thin_border
                c.alignment = center_align
            current_row += 1
        
        current_row += 1 # 코트 간 간격

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="vertical_schedule_{meet.date}.xlsx"'
    wb.save(response)
    return response

# 19. 경기 일정 엑셀 다운로드 (양면 가로형 - 3코트 / 2코트) - 사장님 전용
@login_required
def export_schedule_horizontal(request, meet_id):
    if not is_manager(request.user):
        return redirect('dashboard')

    meet = get_object_or_404(MonthlyMeet, id=meet_id)
    
    # 1. 엑셀 워크북 생성
    wb = Workbook()
    ws = wb.active
    ws.title = f"{meet.date} 경기 일정"

    # 스타일 설정
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True, size=12)
    center_align = Alignment(horizontal='center', vertical='center')

    # 2. 제목 행 설정
    ws.merge_cells('A1:E1')
    ws['A1'] = '3 코트'
    ws.merge_cells('F1:J1')
    ws['F1'] = '2 코트'
    
    for cell in [ws['A1'], ws['F1']]:
        cell.font = Font(bold=True, size=14)
        cell.alignment = center_align

    # 3. 헤더 행 설정
    headers = ['순서', '이름', 'vs', '이름', '비고']
    for col, val in enumerate(headers, 1): # 3코트
        cell = ws.cell(row=2, column=col, value=val)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    for col, val in enumerate(headers, 6): # 2코트
        cell = ws.cell(row=2, column=col, value=val)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # 4. 데이터 영역 채우기 (핸디캡 로직 적용)
    matches_3 = list(Match.objects.filter(meet=meet, court=3).order_by('id'))
    matches_2 = list(Match.objects.filter(meet=meet, court=2).order_by('id'))
    max_len = max(len(matches_3), len(matches_2), 15)
    
    tier_weight = {'A': 3, 'B': 2, 'C': 1}

    for i in range(max_len):
        row_num = i + 3
        
        # --- [좌측: 3 코트] ---
        ws.cell(row=row_num, column=1, value=i+1).border = thin_border
        if i < len(matches_3):
            m = matches_3[i]
            p1_name, p2_name = m.player1.name, m.player2.name
            h = m.applied_handicap
            if h > 0:
                if tier_weight[m.player1.group] < tier_weight[m.player2.group]:
                    p1_name += f" (+{h})"
                elif tier_weight[m.player1.group] > tier_weight[m.player2.group]:
                    p2_name += f" (+{h})"
                else: # 티어 같을 때
                    if m.player1.gender == 'F' and m.player2.gender == 'M': p1_name += f" (+{h})"
                    else: p2_name += f" (+{h})"
            
            ws.cell(row=row_num, column=2, value=p1_name).border = thin_border
            ws.cell(row=row_num, column=3, value='vs').border = thin_border
            ws.cell(row=row_num, column=4, value=p2_name).border = thin_border
        else:
            for c in range(2, 5): ws.cell(row=row_num, column=c).border = thin_border
        ws.cell(row=row_num, column=5).border = thin_border # 비고

        # --- [우측: 2 코트] ---
        ws.cell(row=row_num, column=6, value=i+1).border = thin_border
        if i < len(matches_2):
            m = matches_2[i]
            p1_name, p2_name = m.player1.name, m.player2.name
            h = m.applied_handicap
            if h > 0:
                if tier_weight[m.player1.group] < tier_weight[m.player2.group]:
                    p1_name += f" (+{h})"
                elif tier_weight[m.player1.group] > tier_weight[m.player2.group]:
                    p2_name += f" (+{h})"
                else: # 티어 같을 때
                    if m.player1.gender == 'F' and m.player2.gender == 'M': p1_name += f" (+{h})"
                    else: p2_name += f" (+{h})"

            ws.cell(row=row_num, column=7, value=p1_name).border = thin_border
            ws.cell(row=row_num, column=8, value='vs').border = thin_border
            ws.cell(row=row_num, column=9, value=p2_name).border = thin_border
        else:
            for c in range(7, 10): ws.cell(row=row_num, column=c).border = thin_border
        ws.cell(row=row_num, column=10).border = thin_border # 비고

        # 행 전체 중앙 정렬
        for c in range(1, 11):
            ws.cell(row=row_num, column=c).alignment = center_align

    # 5. 응답 반환
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="schedule_horizontal_{meet.date}.xlsx"'
    wb.save(response)
    return response