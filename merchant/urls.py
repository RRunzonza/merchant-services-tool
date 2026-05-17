from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload_view, name='upload'),
    path('choose/', views.choose_analytics_view, name='choose_analytics'),
    path('select-sheet/', views.select_sheet_view, name='select_sheet'),

    # Daily Performance Analytics
    path('zwg/', views.zwg_performance_view, name='zwg_performance'),
    path('usd/', views.usd_performance_view, name='usd_performance'),
    path('pos/', views.pos_statements_view, name='pos_statements'),

    # Periodic Data Analytics
    path('periodic/', views.periodic_analytics_view, name='periodic_analytics'),
    path('idle/', views.idle_terminals_view, name='idle_terminals'),
    path('top25/zwg/', views.zwg_top25_view, name='zwg_top25'),
    path('top25/usd/', views.usd_top25_view, name='usd_top25'),

    # Downloads — Daily
    path('download/zwg-report/',         views.download_zwg_report,        name='download_zwg_report'),
    path('download/usd-report/',         views.download_usd_report,        name='download_usd_report'),
    path('download/zwg-gainers/',        views.download_zwg_gainers,       name='download_zwg_gainers'),
    path('download/usd-gainers/',        views.download_usd_gainers,       name='download_usd_gainers'),
    path('download/zwg-top25/',          views.download_zwg_top25,         name='download_zwg_top25'),
    path('download/usd-top25/',          views.download_usd_top25,         name='download_usd_top25'),
    path('download/pos-statement/',      views.download_pos_statement,     name='download_pos_statement'),
    path('download/champions-zinara/',   views.download_champions_zinara,  name='download_champions_zinara'),
    path('download/champions-insurance/', views.download_champions_insurance, name='download_champions_insurance'),

    # Downloads — Periodic
    path('download/idle-total/',         views.download_idle_total,        name='download_idle_total'),
    path('download/idle-unit/',          views.download_idle_unit,         name='download_idle_unit'),
    path('download/zwg-top25-weekly/',   views.download_zwg_top25_weekly,  name='download_zwg_top25_weekly'),
    path('download/usd-top25-weekly/',   views.download_usd_top25_weekly,  name='download_usd_top25_weekly'),
    path('merchant-performance/',        views.merchant_performance_view,  name='merchant_performance'),
    path('download/merchant-perf/',      views.download_merchant_perf_report, name='download_merchant_perf_report'),
    path('download/merchant-idle/',      views.download_merchant_idle,     name='download_merchant_idle'),
    path('bu-performance/',              views.bu_performance_view,        name='bu_performance'),
    path('download/bu-perf/',            views.download_bu_perf_report,    name='download_bu_perf_report'),
    path('download/bu-idle/',            views.download_bu_idle,           name='download_bu_idle'),
    path('sector-performance/',           views.sector_performance_view,     name='sector_performance'),
    path('download/sector-combined/',    views.download_sector_combined,    name='download_sector_combined'),
    path('download/sector-detailed/',    views.download_sector_detailed,    name='download_sector_detailed'),

    path('auto-reports/',              views.auto_reports_view,    name='auto_reports'),
    path('auto-reports/proceed/',      views.auto_reports_proceed, name='auto_reports_proceed'),
    path('download/auto-zwg/',         views.download_auto_zwg,    name='download_auto_zwg'),
    path('download/auto-usd/',         views.download_auto_usd,    name='download_auto_usd'),

    path('logout/', views.logout_view, name='logout'),
]
