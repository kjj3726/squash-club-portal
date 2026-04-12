from django.contrib import admin
from .models import Profile, MonthlyMeet, Match

@admin.register(Profile)
class ProfileAdmin(admin.ModelAdmin):
    list_display = ('name', 'group', 'gender', 'user')
    list_filter = ('group', 'gender')

@admin.register(MonthlyMeet)
class MonthlyMeetAdmin(admin.ModelAdmin):
    list_display = ('title', 'date')

@admin.register(Match)
class MatchAdmin(admin.ModelAdmin):
    list_display = ('meet', 'court', 'player1', 'player2', 'is_completed')
    list_filter = ('meet', 'court', 'is_completed')