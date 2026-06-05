from django.contrib.auth import get_user_model
from django.urls import reverse
from django.test import Client
from core.models import UserProfile

User = get_user_model()
username='qc_test_user'
user, created = User.objects.get_or_create(username=username, defaults={'password':'!'})
profile, _ = UserProfile.objects.get_or_create(user=user)
profile.role = 'qc'
profile.save(update_fields=['role'])
print('profile role', profile.role)
print('can_view_sku_master_review_queue', profile.can_view_sku_master_review_queue())
print('can_view_approval_queue', profile.can_view_approval_queue())
print('can_approve_sku_master_review', profile.can_approve_sku_master_review())
client = Client()
client.force_login(user)
resp = client.get(reverse('qc:master_review'))
print('status', resp.status_code)
print('redirect chain', resp.redirect_chain)
print('content snippet', resp.content[:400])
