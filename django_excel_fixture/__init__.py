from django.core.serializers import register_serializer
register_serializer("xlsx", 'django_excel_fixture.serializers.xlsx_serializer')

