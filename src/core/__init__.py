from django.core.serializers import register_serializer
register_serializer("xlsx", 'expenses.core.serializers.excel')
