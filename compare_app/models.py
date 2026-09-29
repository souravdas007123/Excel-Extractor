from django.db import models
from django.contrib.auth.models import User

class UserProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    is_pro = models.BooleanField(default=False)  # Har naya user pehle Free hoga
    total_rows_processed = models.IntegerField(default=0) # Data tracking ke liye
    
    def __str__(self):
        return self.user.username