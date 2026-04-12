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

]