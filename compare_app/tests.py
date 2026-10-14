"""
Django view tests. Run everything with:   python manage.py test
(Pure data-cleaning tests live in test_cleaning.py and need no database.)
"""
from django.contrib.auth.models import User
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.urls import reverse

from .models import UserProfile

PROCESS_FLAGS = {
    'mode': 'single',
    'columns': 'Name, Phone, State',
    'remove_duplicates': 'yes',
    'remove_blanks': 'yes',
    'remove_invalid_phones': 'yes',
    'auto_fix': 'yes',
}


def csv_upload(text, name='leads.csv'):
    return SimpleUploadedFile(name, text.encode('utf-8'), content_type='text/csv')


class PageTests(TestCase):
    def test_public_pages_load(self):
        for name in ('home', 'login', 'register', 'upgrade', 'app'):
            with self.subTest(page=name):
                self.assertEqual(self.client.get(reverse(name)).status_code, 200)


class AccountTests(TestCase):
    def test_register_creates_user_and_profile(self):
        response = self.client.post(reverse('register'), {
            'username': 'newuser',
            'email': 'new@example.com',
            'password1': 'Str0ng-Passw0rd!x',
            'password2': 'Str0ng-Passw0rd!x',
        })
        self.assertRedirects(response, reverse('app'))
        user = User.objects.get(username='newuser')
        self.assertTrue(UserProfile.objects.filter(user=user).exists())

    def test_duplicate_email_is_rejected(self):
        User.objects.create_user('first', 'dup@example.com', 'x')
        response = self.client.post(reverse('register'), {
            'username': 'second',
            'email': 'DUP@example.com',
            'password1': 'Str0ng-Passw0rd!x',
            'password2': 'Str0ng-Passw0rd!x',
        })
        self.assertEqual(response.status_code, 200)
        self.assertFalse(User.objects.filter(username='second').exists())


class ProcessingTests(TestCase):
    def test_csv_single_mode_removes_duplicate(self):
        csv = (
            "Name,Phone,State\n"
            "rahul sharma,+91 98765 43210,delhi\n"
            "Rahul Sharma,9876543210,Delhi\n"
        )
        response = self.client.post(
            reverse('app'), {**PROCESS_FLAGS, 'file_single': csv_upload(csv)}
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.context['single_success'])
        self.assertEqual(response.context['s_duplicate_count'], 1)
        self.assertEqual(response.context['s_fresh_count'], 1)

    def test_free_plan_row_limit_shows_upgrade_card(self):
        rows = "\n".join(f"User {i},9{i:09d},Goa" for i in range(1001))
        csv = "Name,Phone,State\n" + rows + "\n"
        response = self.client.post(
            reverse('app'), {**PROCESS_FLAGS, 'file_single': csv_upload(csv)}
        )
        self.assertIn('limit_error', response.context)
        self.assertNotIn('single_success', response.context)

    def test_missing_column_gives_clear_error(self):
        csv = "Name,Phone\nA,9999999999\n"
        response = self.client.post(
            reverse('app'), {**PROCESS_FLAGS, 'file_single': csv_upload(csv)}
        )
        self.assertIn('Columns not found', response.context['error'])

    def test_wrong_file_type_is_rejected(self):
        bad = SimpleUploadedFile('notes.txt', b'hello', content_type='text/plain')
        response = self.client.post(reverse('app'), {**PROCESS_FLAGS, 'file_single': bad})
        self.assertIn('allowed', response.context['error'])

    def test_usage_counter_increases_for_logged_in_user(self):
        user = User.objects.create_user('counter', 'c@example.com', 'Str0ng-Passw0rd!x')
        self.client.force_login(user)
        csv = "Name,Phone,State\nA One,9111111111,UP\nB Two,9222222222,UP\n"
        self.client.post(reverse('app'), {**PROCESS_FLAGS, 'file_single': csv_upload(csv)})
        self.assertEqual(UserProfile.objects.get(user=user).total_rows_processed, 2)


class DownloadTests(TestCase):
    def test_unknown_token_is_404(self):
        self.assertEqual(self.client.get(reverse('download', args=['0' * 32])).status_code, 404)

    def test_malformed_token_is_404(self):
        self.assertEqual(self.client.get(reverse('download', args=['not-a-valid-token'])).status_code, 404)
