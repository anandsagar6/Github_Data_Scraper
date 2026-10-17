"""Hardening pass: regressions for the seven live bugs, lock-file safety,
transient-failure policy, output validation and legacy migration.
"""

import json
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook       # noqa: E402

import config as C                                 # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (Candidate, TransientFailureLog, XlsxReader)   # noqa: E402

FIXTURES = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "fixtures")
# Real Instant Data Scraper exports, kept in the repo so the suite does not
# depend on anything in ~/Downloads.
REPO_SEARCH = os.path.join(FIXTURES, "github_repo_search.xlsx")
REPO_SEARCH_2 = os.path.join(FIXTURES, "github_repo_search_2.xlsx")
USER_SEARCH = os.path.join(FIXTURES, "github_sample_3.xlsx")


# ==================================================== 8. live-bug regressions
class TestLiveBugRegressions(unittest.TestCase):
    """One test per bug found during live running."""

    def test_bug1_email_followed_by_sentence_punctuation(self):
        # The address used to match NOTHING when a full stop followed it.
        self.assertEqual(student_scraper.extract_emails("write to me@site.com."),
                         ["me@site.com"])
        self.assertEqual(
            student_scraper.extract_emails("Contact: foo.bar+x@sub.site.co.in."),
            ["foo.bar+x@sub.site.co.in"])
        self.assertEqual(student_scraper.extract_emails("(reach anil@gmail.com)"),
                         ["anil@gmail.com"])
        self.assertEqual(student_scraper.extract_emails("mail a@b.org, then call"),
                         ["a@b.org"])

    def test_bug2_whatsapp_number_does_not_claim_phone_column(self):
        from student_scraper import StudentMerger
        # A source whose ONLY number is labelled WhatsApp must not supply
        # the plain Phone Number while another source states a real phone.
        self.assertEqual(
            StudentMerger._one_phone("WhatsApp: 9123456789"), "")
        self.assertEqual(
            StudentMerger._one_phone("Phone: 9876543210"), "9876543210")
        # Both present: the labelled phone wins, WhatsApp still detected.
        text = "Phone: 9876543210\nWhatsApp: 9123456789"
        self.assertEqual(StudentMerger._one_phone(text), "9876543210")
        self.assertEqual(student_scraper.extract_whatsapp(text), "9123456789")

    def test_bug3_internship_does_not_override_total_experience(self):
        resume = ("EXPERIENCE\n"
                  "Software Intern at Infosys, 6 months\n"
                  "2 years experience building web applications")
        self.assertEqual(student_scraper.extract_experience(resume), 24)
        self.assertEqual(student_scraper.format_experience(student_scraper.extract_experience(resume)),
                         "2 years")
        # An internship alone is still reported.
        self.assertEqual(
            student_scraper.extract_experience("Software Intern at Infosys, 6 months"), 6)

    def test_bug4_bare_ai_is_not_a_specialization(self):
        prose = ("Backend Engineering - AI | Exploring Agentic AI, "
                 "AI-assisted development and AI-enabled platforms")
        self.assertEqual(student_scraper.normalize_specialization(prose), "")
        self.assertEqual(
            student_scraper.normalize_specialization("3rd year student. Building AI apps"),
            "")
        # A genuine branch statement still works.
        self.assertEqual(student_scraper.normalize_specialization("pursuing B.Tech in AI"),
                         "Artificial Intelligence")
        self.assertEqual(student_scraper.normalize_specialization("AI/ML branch"),
                         "Artificial Intelligence & Machine Learning")

    def test_bug5_single_letters_in_prose_are_not_languages(self):
        for text in ("Backend Engineering | R&D team",
                     "I can C the problem",
                     "Rated R for content",
                     "Plan B and plan C"):
            found = student_scraper.find_skills(text, student_scraper.PROGRAMMING_LANGUAGES)
            self.assertNotIn("R", found, text)
            self.assertNotIn("C", found, text)
        # Real declarations still count.
        self.assertIn("R", student_scraper.find_skills("Languages: Python, R, SQL",
                                         student_scraper.PROGRAMMING_LANGUAGES))
        self.assertIn("C", student_scraper.find_skills("C and C++ developer",
                                         student_scraper.PROGRAMMING_LANGUAGES))

    def test_bug6_lowercase_github_org_mention(self):
        # GitHub renders "@NatwestGroup" as "@natwestgroup".
        self.assertEqual(
            student_scraper.extract_company_role("Lead Engineer @natwestgroup | Ex-Amazon"),
            ("Lead Engineer", "natwestgroup"))
        self.assertEqual(
            student_scraper.extract_company_role("Web & App Dev Lead @gdsc LNCTS"),
            ("Web & App Dev Lead", "gdsc LNCTS"))
        # "at" still demands a capitalised company.
        self.assertEqual(student_scraper.extract_company_role("engineer at the office"),
                         ("", ""))

    def test_bug7_repository_export_invents_nothing(self):
        """A repo search has no person name and no location."""
        for path in (REPO_SEARCH, REPO_SEARCH_2):
            rows = XlsxReader().read(path, "fresher")
            self.assertTrue(rows, path)
            for row in rows:
                self.assertEqual(row.xlsx_name, "", row.username)
                self.assertEqual(row.xlsx_location, "", row.username)
                self.assertTrue(row.username)
        # A USER export, by contrast, does carry real names.
        users = XlsxReader().read(USER_SEARCH, "lnct")
        self.assertTrue(any(row.xlsx_name for row in users))


# ============================================== 8. additional required cases
class TestRequiredPatterns(unittest.TestCase):
    def test_plus91_phone_formats(self):
        for text in ("Phone +91 9876543210", "call +91-98765-43210",
                     "+919876543210", "tel: 91 9876543210",
                     "mobile +91 98765 43210"):
            found = student_scraper.extract_phones(text)
            self.assertTrue(found, text)
            self.assertEqual(found[0][0], "9876543210", text)

    def test_pin_needs_address_context(self):
        self.assertEqual(student_scraper.extract_pin("Address: Bhopal, PIN 462001"),
                         "462001")
        self.assertEqual(student_scraper.extract_pin("Postal Code 110001"), "110001")
        self.assertEqual(student_scraper.extract_pin("462001"), "")
        self.assertEqual(student_scraper.extract_pin("build 462001 passed"), "")
        self.assertEqual(student_scraper.extract_pin("PIN 062001"), "")

    def test_short_year_needs_education_context(self):
        self.assertEqual(student_scraper.extract_graduation_year("LNCT '26"), "2026")
        self.assertEqual(
            student_scraper.extract_graduation_year("B.Tech student, batch '27"), "2027")
        self.assertEqual(student_scraper.extract_graduation_year("he said 'never'"), "")
        self.assertEqual(student_scraper.extract_graduation_year("rock 'n' roll '26 tour"),
                         "")

    def test_year_range_gives_the_end_year(self):
        self.assertEqual(student_scraper.extract_graduation_year("B.Tech CSE 2023-2027"),
                         "2027")
        self.assertEqual(
            student_scraper.extract_graduation_year("College 2021 to 2025"), "2025")
        # Outside an education context a range means nothing.
        self.assertEqual(
            student_scraper.extract_graduation_year("server uptime 2019-2024"), "")

    def test_linkedin_personal_versus_company(self):
        self.assertEqual(
            student_scraper.extract_linkedin("https://www.linkedin.com/in/anil-p/?utm_a=1"),
            "https://www.linkedin.com/in/anil-p")
        self.assertEqual(
            student_scraper.extract_linkedin("https://in.linkedin.com/in/foo"),
            "https://www.linkedin.com/in/foo")
        for company in ("https://www.linkedin.com/company/acme",
                        "https://linkedin.com/school/vit",
                        "https://www.linkedin.com/showcase/x"):
            self.assertEqual(student_scraper.extract_linkedin(company), "", company)

    def test_public_resume_pdf_url(self):
        self.assertEqual(
            student_scraper.extract_resume_urls("Download my resume "
                                  "https://site.com/files/My_Resume.pdf"),
            ["https://site.com/files/My_Resume.pdf"])
        self.assertIn("https://site.com/cv.pdf",
                      student_scraper.extract_resume_urls("CV: https://site.com/cv.pdf"))
        self.assertEqual(student_scraper.extract_resume_urls("https://site.com/index.html"),
                         [])

    def test_template_email_rejection(self):
        for bad in ("your_email@example.com", "example@example.com",
                    "test@test.com", "name@email.com",
                    "bot@users.noreply.github.com", "no-reply@github.com",
                    "contact@example.org"):
            self.assertEqual(student_scraper.extract_emails(bad), [], bad)


# ================================================== 7. XLSX source row check
class TestXlsxSourceRows(unittest.TestCase):
    def test_real_sample_files_parse_without_fixed_positions(self):
        for path in (REPO_SEARCH, REPO_SEARCH_2, USER_SEARCH):
            rows = XlsxReader().read(path, "fresher")
            self.assertTrue(rows, path)
            for row in rows:
                self.assertTrue(row.username)
                self.assertEqual(row.profile_url,
                                 "https://github.com/%s" % row.username)
                self.assertTrue(row.raw_cells, "raw cells must be preserved")

    def test_user_export_reconstructs_split_highlight_text(self):
        """GitHub splits a highlighted term into its own column."""
        rows = {row.username: row
                for row in XlsxReader().read(USER_SEARCH, "lnct")}
        danish = rows["letscodedanish"]
        self.assertEqual(danish.xlsx_name, "Danish Khan")
        self.assertEqual(danish.profile_url,
                         "https://github.com/letscodedanish")
        # "3rd year student at  BHOPAL." + "LNCT" rejoined.
        self.assertIn("3rd year student at", danish.xlsx_bio)
        self.assertIn("LNCT", danish.xlsx_bio)
        # The login is never repeated as biography.
        self.assertNotIn("letscodedanish", danish.xlsx_bio)

        gourav = rows["GouravRusiya30"]
        self.assertEqual(gourav.xlsx_name, "Gourav Rusiya")
        self.assertEqual(gourav.xlsx_location, "Bangalore")
        self.assertIn("Lead Engineer", gourav.xlsx_bio)
        # A follower/repo count is not a location.
        self.assertNotIn("29", gourav.xlsx_location)

    def test_user_export_keeps_every_useful_cell_for_temporary_txt(self):
        rows = {row.username: row
                for row in XlsxReader().read(USER_SEARCH, "lnct")}
        cells = rows["khush1yaaar"].raw_cells
        self.assertTrue(any("khush1yaaar" in cell for cell in cells))
        self.assertTrue(any("Khushi" in cell for cell in cells))
        self.assertTrue(any("github.com" in cell for cell in cells))

    def test_shuffled_columns_give_the_same_answer(self):
        directory = tempfile.mkdtemp(prefix="shuffle_")
        try:
            cells = ["Danish Khan", "https://github.com/letscodedanish",
                     "3rd year student at", "LNCT",
                     "https://github.com/letscodedanish.png?size=40"]
            results = []
            for order in ([0, 1, 2, 3, 4], [4, 2, 1, 3, 0], [3, 0, 4, 1, 2]):
                path = os.path.join(directory, "e%d.xlsx" % order[0])
                book = Workbook()
                sheet = book.active
                sheet.append(["c%d" % n for n in range(5)])
                sheet.append([cells[n] for n in order])
                book.save(path)
                row = XlsxReader().read(path, "LNCT")[0]
                results.append((row.username, row.profile_url))
            self.assertEqual(len(set(results)), 1, results)
            self.assertEqual(results[0][0], "letscodedanish")
        finally:
            shutil.rmtree(directory, ignore_errors=True)


# ================================================ 2. lock-file safety
class TestLockFiles(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="locks_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _lock(self, target):
        path = os.path.join(self.dir, ".~lock.%s#" % target)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(",anand,host1,24.09.2026 17:54,file:///x;")
        return path

    def test_active_lock_is_detected_and_described(self):
        Workbook().save(os.path.join(self.dir, "student_data_2.xlsx"))
        self._lock("student_data_2.xlsx")
        locks = student_scraper.find_lock_files(self.dir)
        self.assertEqual(len(locks), 1)
        self.assertFalse(locks[0]["orphan"])
        self.assertEqual(locks[0]["user"], "anand")
        text = student_scraper.describe_locks(locks)[0]
        self.assertIn("student_data_2.xlsx", text)
        self.assertIn("open in LibreOffice", text)

    def test_orphan_lock_is_flagged_but_not_active(self):
        self._lock("student_data_9.xlsx")          # workbook absent
        locks = student_scraper.find_lock_files(self.dir)
        self.assertTrue(locks[0]["orphan"])
        self.assertIn("stale", student_scraper.describe_locks(locks)[0])

    def test_lock_file_is_never_removed(self):
        Workbook().save(os.path.join(self.dir, "student_data_1.xlsx"))
        path = self._lock("student_data_1.xlsx")
        student_scraper.find_lock_files(self.dir)
        student_scraper.describe_locks(student_scraper.find_lock_files(self.dir))
        self.assertTrue(os.path.exists(path),
                        "a lock file must never be deleted automatically")

    def test_writing_refuses_while_the_target_is_locked(self):
        """The guard moved from the scraper to the export; it still refuses."""
        Workbook().save(os.path.join(self.dir, "student_data.xlsx"))
        self._lock("student_data.xlsx")
        with self.assertRaises(student_scraper.WorkbookLocked):
            student_scraper.guard_locks(self.dir)

    def test_a_stale_lock_is_reported_and_does_not_block(self):
        self._lock("student_data.xlsx")          # workbook absent -> stale
        messages = []
        stale = student_scraper.guard_locks(self.dir, say=messages.append)
        self.assertEqual(len(stale), 1)
        self.assertTrue(any("stale" in line for line in messages), messages)


# ================================================ 6. output validation
class TestOutputValidation(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="validate_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _record(self, number):
        record = {column: "" for column in C.FINAL_COLUMNS}
        record["GitHub Profile"] = "https://github.com/person%d" % number
        record["github"] = record["GitHub Profile"]
        record["Country"] = "India"
        record["country"] = "India"
        return record

    def _write(self, count):
        path = os.path.join(self.dir, "student_data.xlsx")
        student_scraper.write_final([self._record(number)
                       for number in range(1, count + 1)], path)
        return path

    def test_clean_output_passes(self):
        path = self._write(5)
        ok, problems, summary = student_scraper.validate_output(path)
        self.assertTrue(ok, problems)
        self.assertEqual(summary, [("student_data.xlsx", 65, 5, "PASS")])

    def test_alias_mismatch_is_caught(self):
        path = self._write(1)
        book = load_workbook(path)
        sheet = book.active
        column = C.FINAL_COLUMNS.index("country") + 1
        sheet.cell(row=2, column=column, value="Nepal")     # canonical: India
        book.save(path)
        book.close()
        ok, problems, _summary = student_scraper.validate_output(path)
        self.assertFalse(ok)
        self.assertTrue(any("alias" in problem for problem in problems),
                        problems)

    def test_duplicate_username_in_the_workbook_is_caught(self):
        path = self._write(2)
        book = load_workbook(path)
        sheet = book.active
        column = C.FINAL_COLUMNS.index("GitHub Profile") + 1
        alias = C.FINAL_COLUMNS.index("github") + 1
        value = sheet.cell(row=2, column=column).value
        sheet.cell(row=3, column=column, value=value)
        sheet.cell(row=3, column=alias, value=value)
        book.save(path)
        book.close()
        ok, problems, _summary = student_scraper.validate_output(path)
        self.assertFalse(ok)
        self.assertTrue(any("duplicate" in problem for problem in problems),
                        problems)

    def test_the_row_maximum_is_the_xlsx_limit_not_a_hundred(self):
        """There is no 100-row rule any more; 150 rows is simply valid."""
        path = self._write(150)
        ok, problems, summary = student_scraper.validate_output(path)
        self.assertTrue(ok, problems)
        self.assertEqual(summary[0][2], 150)

    def test_legacy_schema_is_reported(self):
        book = Workbook()
        book.active.append(student_scraper.LEGACY_COLUMNS)
        book.save(os.path.join(self.dir, "student_data_1.xlsx"))
        book.close()
        ok, problems, _summary = student_scraper.validate_output(
            os.path.join(self.dir, "student_data.xlsx"))
        self.assertFalse(ok)
        self.assertTrue(any("student_data_1.xlsx" in problem
                            for problem in problems), problems)


# ================================================ 1. legacy migration
class TestLegacyMigration(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="migrate_")
        self.backup = tempfile.mkdtemp(prefix="migrate_backup_")
        self.legacy = os.path.join(self.dir, "student_data_1.xlsx")
        book = Workbook()
        sheet = book.active
        sheet.append(["ID", "Keyword", "Username", "Name", "Email", "Phone",
                      "Location", "Bio", "Links", "Profile URL", "Repo URL",
                      "Repo Name", "Stars", "Source Row", "Source File",
                      "Scraped At"])
        sheet.append([1, "B.Tech Student", "person1", "Person One",
                      "person1@gmail.com", "9876543210",
                      "Bhopal, Madhya Pradesh",
                      "B.Tech CSE student, batch 2023-2027",
                      "https://www.linkedin.com/in/person-one", 
                      "https://github.com/person1", "", "", "",
                      "Person One | person1", "btech-student.xlsx",
                      "2026-09-24 17:10:32"])
        sheet.append([2, "B.Tech Student", "person2", "Person Two",
                      "Not found", "Not found", "Not found", "Some bio",
                      "None", "https://github.com/person2", "", "", "",
                      "Person Two", "btech-student.xlsx",
                      "2026-09-24 17:10:34"])
        book.save(self.legacy)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)
        shutil.rmtree(self.backup, ignore_errors=True)

    def test_reading_changes_nothing(self):
        """Migration is read-only on the source; --apply writes a store."""
        before = sorted(os.listdir(self.dir))
        size = os.path.getsize(self.legacy)
        records, problems = student_scraper.read_workbook_records(self.legacy)
        self.assertEqual(len(records), 2)
        self.assertEqual(problems, [])
        self.assertEqual(sorted(os.listdir(self.dir)), before)
        self.assertEqual(os.path.getsize(self.legacy), size)

    def test_migration_preserves_every_row_and_the_original(self):
        records, _problems = student_scraper.read_workbook_records(self.legacy)
        store = student_scraper.open_store(os.path.join(self.backup,
                                                     "legacy.db"))
        self.addCleanup(store.close)
        for record in records:
            store.save_student(record, campaign_id="legacy-import")
        self.assertEqual(store.count(), 2)
        self.assertTrue(os.path.exists(self.legacy),
                        "the source workbook must be PRESERVED")

    def test_field_mapping(self):
        records, _problems = student_scraper.read_workbook_records(self.legacy)
        first = records[0]
        self.assertEqual(first["GitHub Profile"],
                         "https://github.com/person1")
        self.assertEqual(first["github"], first["GitHub Profile"])
        self.assertEqual(first["Student Name"], "Person One")
        self.assertEqual(first["Email Address"], "person1@gmail.com")
        self.assertEqual(first["Phone Number"], "9876543210")
        self.assertEqual(first["Current City"], "Bhopal")
        self.assertEqual(first["State"], "Madhya Pradesh")
        self.assertEqual(first["Country"], "India")
        self.assertEqual(first["LinkedIn Profile"],
                         "https://www.linkedin.com/in/person-one")
        self.assertEqual(first["Source URL"], first["GitHub Profile"])
        self.assertEqual(first["Data Collection Date"], "2026-09-24")
        self.assertIn("Migrated", first["Remarks / Notes"])
        for alias, canonical in C.ALIAS_COLUMNS.items():
            self.assertEqual(first[alias], first[canonical])

    def test_sentinel_values_become_blank_not_literal(self):
        records, _problems = student_scraper.read_workbook_records(self.legacy)
        second = records[1]
        self.assertEqual(second["Email Address"], "")
        self.assertEqual(second["Phone Number"], "")
        self.assertEqual(second["Current City"], "")
        self.assertNotIn("Not found",
                         "".join(str(value) for value in second.values()))

    def test_no_duplicate_created_when_person_already_migrated(self):
        records, _problems = student_scraper.read_workbook_records(self.legacy)
        store = student_scraper.open_store(os.path.join(self.backup,
                                                     "legacy.db"))
        self.addCleanup(store.close)
        for _pass in range(2):
            for record in records:
                store.save_student(record, campaign_id="legacy-import")
        self.assertEqual(store.count(), 2)


# ================================================ 4. transient failures
class TestTransientFailures(unittest.TestCase):
    def setUp(self):
        handle, self.path = tempfile.mkstemp(suffix=".json")
        os.close(handle)
        os.remove(self.path)
        self.candidate = Candidate(
            source_row_number=137, source_file="btech-student.xlsx",
            keyword="B.Tech Student", username="xyz",
            profile_url="https://github.com/xyz")

    def tearDown(self):
        for suffix in ("", ".tmp"):
            if os.path.exists(self.path + suffix):
                os.remove(self.path + suffix)

    def test_records_every_required_detail(self):
        log = TransientFailureLog(path=self.path)
        entry = log.record(self.candidate, "navigation timeout")
        for key in ("username", "github_url", "keyword", "source_xlsx",
                    "source_row", "reason", "attempts", "last_attempt"):
            self.assertIn(key, entry)
        self.assertEqual(entry["username"], "xyz")
        self.assertEqual(entry["github_url"], "https://github.com/xyz")
        self.assertEqual(entry["keyword"], "B.Tech Student")
        self.assertEqual(entry["source_xlsx"], "btech-student.xlsx")
        self.assertEqual(entry["source_row"], 137)
        self.assertEqual(entry["reason"], "navigation timeout")
        self.assertEqual(entry["attempts"], 1)

    def test_attempts_accumulate_and_persist(self):
        log = TransientFailureLog(path=self.path)
        for _ in range(3):
            log.record(self.candidate, "timeout")
        self.assertEqual(TransientFailureLog(path=self.path).attempts("xyz"), 3)
        self.assertEqual(TransientFailureLog(path=self.path).attempts("XYZ"), 3)

    def test_blocking_only_after_retries_are_exhausted(self):
        log = TransientFailureLog(path=self.path)
        log.record(self.candidate, "timeout")
        self.assertEqual(log.blocking(), [])
        for _ in range(C.PROFILE_ATTEMPTS_MAX - 1):
            log.record(self.candidate, "timeout")
        self.assertEqual(len(log.blocking()), 1)

    def test_success_clears_the_entry_no_blacklist(self):
        log = TransientFailureLog(path=self.path)
        log.record(self.candidate, "timeout")
        log.clear("xyz")
        self.assertEqual(log.attempts("xyz"), 0)
        self.assertEqual(TransientFailureLog(path=self.path).blocking(), [])

    def test_file_is_valid_json_on_disk(self):
        log = TransientFailureLog(path=self.path)
        log.record(self.candidate, "timeout")
        with open(self.path, encoding="utf-8") as handle:
            data = json.load(handle)
        self.assertIn("xyz", data["failures"])



class TestNoConfigLookupHidesADeletedConstant(unittest.TestCase):
    """`getattr(C, "NAME", default)` must not be how a config flag is read.

    WHY THIS EXISTS
    --------------
    `PROFILE_RATE_LIMIT_COOLDOWN` was read as
    `getattr(C, "PROFILE_RATE_LIMIT_COOLDOWN", False)`.  An audit for unused
    constants searched for `C.NAME`, did not see the string form, concluded the
    constant was dead and deleted it -- and the getattr default silently turned
    rate-limit cooldown OFF instead of raising.  One test caught it; nothing
    about the deletion itself did.

    Direct attribute access makes a deleted constant an AttributeError on the
    first run, which is the behaviour you want from a safety flag.
    """

    def test_no_getattr_default_on_the_config_module(self):
        import ast
        import student_scraper as S
        source = open(S.__file__, encoding="utf-8").read()
        offenders = []
        for node in ast.walk(ast.parse(source)):
            if not isinstance(node, ast.Call):
                continue
            if not (isinstance(node.func, ast.Name) and node.func.id == "getattr"):
                continue
            if len(node.args) < 2 or not isinstance(node.args[0], ast.Name):
                continue
            if node.args[0].id not in ("C", "config"):
                continue
            name = (node.args[1].value
                    if isinstance(node.args[1], ast.Constant) else "?")
            offenders.append("line %d: getattr(%s, %r, ...)"
                             % (node.lineno, node.args[0].id, name))
        self.assertEqual(offenders, [],
                         "read config constants directly so that deleting one "
                         "raises instead of silently changing behaviour:\n  "
                         + "\n  ".join(offenders))


if __name__ == "__main__":
    unittest.main(verbosity=2)
