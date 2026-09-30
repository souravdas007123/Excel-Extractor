from datetime import date

from django.contrib.auth.models import User
from django.db import models
from django.db.models.signals import post_save
from django.dispatch import receiver


class UserProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    is_pro = models.BooleanField(default=False)
    pro_expires_on = models.DateField(null=True, blank=True)  # blank = no expiry
    total_rows_processed = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True, null=True)

    def has_active_pro(self):
        if not self.is_pro:
            return False
        if self.pro_expires_on and self.pro_expires_on < date.today():
            return False
        return True

    def __str__(self):
        return self.user.username


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    if created:
        UserProfile.objects.get_or_create(user=instance)