from django.db import models
from django.contrib.auth import get_user_model
from django.utils import timezone
import datetime

User = get_user_model()

class BackupSetting(models.Model):
    FREQUENCY_CHOICES = [
        ('DAILY', 'Daily'),
        ('WEEKLY', 'Weekly'),
        ('MONTHLY', 'Monthly'),
    ]

    backup_enabled = models.BooleanField(default=True, help_text="Enable or disable automated backups.")
    backup_time = models.TimeField(default=datetime.time(20, 5), help_text="Time of day when the backup should run.")
    frequency = models.CharField(max_length=10, choices=FREQUENCY_CHOICES, default='DAILY', help_text="How often backups should run.")
    local_backup_folder = models.CharField(max_length=255, default='backups', help_text="Local directory where backups will be stored.")
    
    cloud_onedrive_folder = models.CharField(max_length=255, blank=True, null=True, help_text="Windows: path to OneDrive sync folder (e.g. C:\\Users\\Name\\OneDrive\\ERP_Backups). Linux/cloud servers with no desktop client: an rclone remote instead, e.g. onedrive:ERP_Backups/CloudVM")
    cloud_gdrive_folder = models.CharField(max_length=255, blank=True, null=True, help_text="Windows: path to Google Drive sync folder (e.g. G:\\My Drive\\ERP_Backups). Linux/cloud servers with no desktop client: an rclone remote instead, e.g. gdrive:ERP_Backups/CloudVM")
    
    keep_daily = models.IntegerField(default=30, help_text="Number of daily backups to retain.")
    keep_weekly = models.IntegerField(default=12, help_text="Number of weekly backups to retain.")
    keep_monthly = models.IntegerField(default=12, help_text="Number of monthly backups to retain.")
    
    include_media = models.BooleanField(default=False, help_text="Include the media folder in the backup.")
    media_cloud_folder = models.CharField(max_length=255, blank=True, null=True, help_text="Optional: send media to a DIFFERENT destination than the database backup (local path or rclone remote, e.g. gdrive:ERP_Backups/CloudVM). Only used when 'Include media' is on. Leave blank to bundle media into the same zip as the database, sent to both OneDrive/Google Drive folders above as usual.")
    include_logs = models.BooleanField(default=False, help_text="Include system logs inside the backup archive.")
    enable_notifications = models.BooleanField(default=True, help_text="Enable notifications on backup success/failure.")
    enable_encryption = models.BooleanField(default=False, help_text="Enable ZIP password encryption.")
    encryption_password = models.CharField(max_length=128, blank=True, null=True, help_text="ZIP protection password.")

    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Backup Setting"
        verbose_name_plural = "Backup Settings"

    @classmethod
    def get_settings(cls):
        obj, created = cls.objects.get_or_create(id=1)
        return obj

    def __str__(self):
        return f"Backup Settings (Enabled: {self.backup_enabled}, Time: {self.backup_time}, Freq: {self.frequency})"


class BackupHistory(models.Model):
    TYPE_CHOICES = [
        ('AUTO', 'Automatic'),
        ('MANUAL', 'Manual'),
    ]
    STATUS_CHOICES = [
        ('PENDING', 'Pending'),
        ('SUCCESS', 'Success'),
        ('FAILED', 'Failed'),
    ]

    backup_type = models.CharField(max_length=10, choices=TYPE_CHOICES, default='AUTO')
    start_time = models.DateTimeField(default=timezone.now)
    finish_time = models.DateTimeField(null=True, blank=True)
    duration_seconds = models.IntegerField(null=True, blank=True, help_text="Duration of the backup operation in seconds.")
    file_name = models.CharField(max_length=255, blank=True, null=True)
    file_size = models.BigIntegerField(null=True, blank=True, help_text="Size of backup file in bytes.")
    backup_location = models.TextField(blank=True, null=True, help_text="Path(s) where the backup is stored.")
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='PENDING')
    error_message = models.TextField(blank=True, null=True)
    sha256_checksum = models.CharField(max_length=64, blank=True, null=True, help_text="SHA-256 hash of the generated backup zip file.")
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, help_text="User who initiated manual backup (null for automatic).")

    class Meta:
        verbose_name = "Backup History"
        verbose_name_plural = "Backup Histories"
        ordering = ['-start_time']

    def __str__(self):
        return f"Backup {self.id} - {self.file_name or 'Pending'} ({self.status})"


class RestoreHistory(models.Model):
    STATUS_CHOICES = [
        ('SUCCESS', 'Success'),
        ('FAILED', 'Failed'),
    ]

    backup = models.ForeignKey(BackupHistory, on_delete=models.SET_NULL, null=True, blank=True)
    timestamp = models.DateTimeField(default=timezone.now)
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='SUCCESS')
    error_message = models.TextField(blank=True, null=True)
    executed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)

    class Meta:
        verbose_name = "Restore History"
        verbose_name_plural = "Restore Histories"
        ordering = ['-timestamp']

    def __str__(self):
        return f"Restore {self.id} on {self.timestamp} - {self.status}"
