from django.contrib import admin
from django.urls import path
from django.contrib.auth import views as auth_views
from matches import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.dashboard, name='dashboard'),
    path('create-meet-matches/', views.create_meet_and_matches, name='create_meet_matches'),
    path('record-score/<int:match_id>/', views.record_score, name='record_score'),
    
    # 로그인 & 로그아웃 & 회원가입
    path('login/', views.custom_login, name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='/'), name='logout'),
    path('signup/', views.signup, name='signup'),
    
    # 관리자용 인원 추가 URL (신규)
    path('add-member/', views.add_member_by_admin, name='add_member_admin'),

    # (신규) 오늘 모임 최종 마감 URL
    path('finalize-meet/<int:meet_id>/', views.finalize_meet, name='finalize_meet'),

    # (신규) 엑셀 CVS 템플릿 다운로드 및 일괄 업로드 URL
    path('download-template/', views.download_member_template, name='download_template'),
    path('upload-bulk/', views.upload_members_bulk, name='upload_bulk'),

    # (신규) 과거 경기 기록 엑셀 CVS 템플릿 다운로드 및 일괄 업로드 URL
    path('download-match-template/', views.download_match_template, name='download_match_template'),
    path('upload-matches-bulk/', views.upload_matches_bulk, name='upload_matches_bulk'),

    # (신규) 사장님만 수정이 가능한 경기 결과 수정 URL
    path('update-match/<int:match_id>/', views.update_match_detail, name='update_match'),

    # (신규) 커스텀 로그아웃 URL
    path('logout/', views.custom_logout, name='logout'),

    # (신규) 비밀번호 변경 URL
    path('change-password/', views.change_password, name='change_password'),

    # (신규) 경기 일정 엑셀 다운로드 URL
    path('export-vertical/<int:meet_id>/', views.export_schedule_vertical, name='export_schedule_vertical'),
    
    # (신규) 경기 일정 엑셀 다운로드 URL (가로형)
    path('export-horizontal/<int:meet_id>/', views.export_schedule_horizontal, name='export_schedule_horizontal'),

    # 🌟 인원 관리용 신규 주소 2개 추가
    path('members/', views.member_management, name='member_management'),
    path('members/update-rank/<int:profile_id>/', views.update_member_rank, name='update_member_rank'),
    
    # (신규) 게스트 인원을 정규 인원으로 승격 URL
    path('members/promote-guest/<int:profile_id>/', views.promote_guest, name='promote_guest'),
    
    # (신규) 경기 결과 엑셀 다운로드 URL
    path('export-results/<int:meet_id>/', views.export_meet_results, name='export_meet_results'),

    # (신규) 경기 결과 수정 및 재조정 URL
    path('meet/<int:meet_id>/rebalance/', views.handle_absentee_and_rebalance, name='handle_absentee_and_rebalance'),

    # (신규) 실시간 점수판 API URL
    path('meet/<int:meet_id>/live-scores/', views.get_live_scores, name='get_live_scores'),

    # (신규) 경기 취소 URL
    path('meet/<int:meet_id>/cancel/', views.cancel_meeting, name='cancel_meeting'),

    # (신규) 핸디캡 수동 수정 URL
    path('meet/<int:meet_id>/handicaps/update/', views.update_handicaps, name='update_handicaps'),

    # 🌟 공지사항 관련 URL 추가
    path('notices/', views.notice_list, name='notice_list'),
    path('notices/save/', views.notice_save, name='notice_save'),
    path('notices/<int:notice_id>/', views.notice_detail, name='notice_detail'),
    path('notices/<int:notice_id>/delete/', views.notice_delete, name='notice_delete'),
    path('notices/<int:notice_id>/add_comment/', views.add_notice_comment, name='add_notice_comment'),
    path('notices/comments/<int:comment_id>/delete/', views.delete_notice_comment, name='delete_notice_comment'),
]