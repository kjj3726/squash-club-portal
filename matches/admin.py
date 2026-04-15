from django.contrib import admin
from django.contrib.admin.models import LogEntry
from .models import AppLog

# 1. Django 관리자 페이지 내 활동 기록(LogEntry)을 볼 수 있도록 등록
@admin.register(LogEntry)
class LogEntryAdmin(admin.ModelAdmin):
    list_display = ('action_time', 'user', 'content_type', 'action_flag', 'change_message')
    list_filter = ('action_flag', 'action_time')
    search_fields = ('user__username', 'change_message')
    readonly_fields = ('action_time', 'user', 'content_type', 'object_id', 'object_repr', 'action_flag', 'change_message')

    def has_add_permission(self, request): return False
    def has_change_permission(self, request, obj=None): return False
    def has_delete_permission(self, request, obj=None): return False

# 2. 서버/시스템 에러 및 경고 로그를 볼 수 있도록 등록
@admin.register(AppLog)
class AppLogAdmin(admin.ModelAdmin):
    list_display = ('level', 'message_preview', 'created_at')
    list_filter = ('level', 'created_at')
    search_fields = ('message',)
    readonly_fields = ('level', 'message', 'created_at')

    def message_preview(self, obj):
        return obj.message[:80] + '...' if len(obj.message) > 80 else obj.message
    message_preview.short_description = '로그 내용'

    def has_add_permission(self, request): return False
    def has_change_permission(self, request, obj=None): return False