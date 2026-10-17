"""End-to-end: sources -> temporary.txt -> final row.

Uses github_profile.txt, a real captured profile that ships with the repo.
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                       # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (Candidate, SourceBundle, StudentMerger)   # noqa: E402
import sandbox                                         # noqa: E402

PROFILE_TXT = os.path.join(C.PROJECT_ROOT, "github_profile.txt")

RESUME = """ANIL PALAGANI
Hyderabad, Telangana | Phone: +91 98765 43210 | anilpalagani9@gmail.com

EDUCATION
B.Tech in Computer Science Engineering, 2021-2025
Lakshmi Narain College of Technology, Bhopal
Affiliated to RGPV University

EXPERIENCE
Software Intern at Infosys, 6 months
2 years experience building web applications

SKILLS
Python, Java, React, Docker, Git, Linux, MySQL
"""

PORTFOLIO = """Anil Palagani
Hi, I'm a final year student.
Location: Hyderabad, Telangana
Contact: anilpalagani9@gmail.com
WhatsApp: 9123456789
Address: Road No 5, Hyderabad, PIN 500081
Interested in Cloud Computing opportunities
"""


class TestMerge(unittest.TestCase):
    def setUp(self):
        with open(PROFILE_TXT, encoding="utf-8") as handle:
            self.profile_text = handle.read()
        self.candidate = Candidate(
            source_row_number=1, source_file="btech-student.xlsx",
            keyword="B.Tech Student", username="Anilyadav9989",
            profile_url="https://github.com/Anilyadav9989",
            xlsx_name="PALAGANI ANIL", xlsx_bio="recently graduate student",
            xlsx_location="", raw_cells=["PALAGANI ANIL", "Anilyadav9989"])

    def _bundle(self, with_portfolio=True, with_resume=True):
        bundle = SourceBundle(candidate=self.candidate)
        bundle.github_ok = True
        bundle.github_text = self.profile_text
        bundle.github_fields = {
            "isUser": True, "nickname": "Anilyadav9989",
            "fullname": "PALAGANI ANIL",
            "bio": "A recently graduate student actively building projects "
                   "and seeking Job opportunities.",
            "location": "Hyderabad, Telangana",
            "email": "anilpalagani9@gmail.com",
            "website": "https://anilpalagani.dev",
            "links": ["https://anilpalagani.dev",
                      "https://www.linkedin.com/in/anil-palagani-a253b1254"],
            "readme": "",
        }
        bundle.github_links = bundle.github_fields["links"]
        if with_portfolio:
            bundle.portfolio_url = "https://anilpalagani.dev"
            bundle.portfolio_text = PORTFOLIO
            bundle.portfolio_links = ["https://anilpalagani.dev/resume.pdf"]
            bundle.portfolio_pages = ["https://anilpalagani.dev"]
        if with_resume:
            bundle.resume_url = "https://anilpalagani.dev/resume.pdf"
            bundle.resume_text = RESUME
        return bundle

    def test_schema_is_exact(self):
        record = StudentMerger().build(self._bundle())
        self.assertEqual(list(record.keys()), C.FINAL_COLUMNS)
        self.assertEqual(len(record), 65)

    def test_core_fields(self):
        record = StudentMerger().build(self._bundle())
        self.assertEqual(record["Student Name"], "PALAGANI ANIL")
        self.assertEqual(record["Email Address"], "anilpalagani9@gmail.com")
        self.assertEqual(record["Phone Number"], "9876543210")
        self.assertEqual(record["WhatsApp Number"], "9123456789")
        self.assertEqual(record["Current City"], "Hyderabad")
        self.assertEqual(record["State"], "Telangana")
        self.assertEqual(record["Country"], "India")
        self.assertEqual(record["PIN Code"], "500081")
        self.assertEqual(record["GitHub Profile"],
                         "https://github.com/Anilyadav9989")
        self.assertEqual(record["LinkedIn Profile"],
                         "https://www.linkedin.com/in/anil-palagani-a253b1254")
        self.assertEqual(record["Portfolio Website"],
                         "https://anilpalagani.dev")
        self.assertEqual(record["Resume URL"],
                         "https://anilpalagani.dev/resume.pdf")

    def test_education(self):
        record = StudentMerger().build(self._bundle())
        self.assertEqual(record["Current Course"], "B.Tech")
        self.assertEqual(record["Current Specialization"],
                         "Computer Science Engineering")
        self.assertEqual(record["Current Education Level"], "Undergraduate")
        self.assertEqual(record["Expected Graduation Year"], "2025")
        self.assertIn("College", record["College Name"])
        # A UG degree fills the UG block and leaves the PG block empty.
        self.assertEqual(record["UG Degree"], "B.Tech")
        self.assertEqual(record["UG Graduation Year"], "2025")
        self.assertEqual(record["PG Degree"], "")

    def test_skills_and_experience(self):
        record = StudentMerger().build(self._bundle())
        for skill in ("Python", "Java"):
            self.assertIn(skill, record["Programming Languages"])
        self.assertIn("React", record["Technologies / Frameworks"])
        self.assertIn("Git", record["Tools / Software Known"])
        self.assertEqual(record["Total Experience"], "2 years")
        self.assertIn("Intern", record["Internship Experience"])

    def test_alias_columns_always_mirror(self):
        record = StudentMerger().build(self._bundle())
        for alias, canonical in C.ALIAS_COLUMNS.items():
            self.assertEqual(record[alias], record[canonical],
                             "%s != %s" % (alias, canonical))
            if record[canonical]:
                self.assertTrue(record[alias], "alias %s left blank" % alias)

    def test_provenance(self):
        record = StudentMerger().build(self._bundle())
        self.assertEqual(record["Data Source"],
                         "GitHub + Portfolio + Resume")
        self.assertEqual(record["Source Platform"], "GitHub")
        self.assertEqual(record["Source URL"],
                         "https://github.com/Anilyadav9989")
        self.assertEqual(record["source_file"], "btech-student.xlsx")
        self.assertRegex(record["Data Collection Date"],
                         r"^\d{4}-\d{2}-\d{2}$")

    def test_data_source_reflects_sources_actually_used(self):
        record = StudentMerger().build(
            self._bundle(with_portfolio=False, with_resume=False))
        self.assertEqual(record["Data Source"], "GitHub")

    def test_validation_statuses(self):
        record = StudentMerger().build(self._bundle())
        self.assertEqual(record["Email Validation Status"], "Valid Format")
        self.assertEqual(record["Phone Validation Status"], "Valid Format")
        self.assertEqual(record["Data Verified Status"],
                         "Verified from public profile")

        bare = self._bundle(with_portfolio=False, with_resume=False)
        bare.github_fields["email"] = ""
        record = StudentMerger().build(bare)
        self.assertEqual(record["Email Validation Status"], "Not Found")
        self.assertEqual(record["Phone Validation Status"], "Not Found")

    def test_graduation_conflict_is_recorded(self):
        bundle = self._bundle()
        bundle.github_fields["bio"] = "B.Tech student, Class of 2026"
        record = StudentMerger().build(bundle)
        self.assertIn("conflict", record["Remarks / Notes"].lower())
        self.assertIn("2026", record["Remarks / Notes"])
        self.assertIn("2025", record["Remarks / Notes"])
        # The resume value wins, and the note says so.
        self.assertEqual(record["Expected Graduation Year"], "2025")

    def test_temporary_file_round_trip(self):
        # temp_store is rank-local now, so the test gives it a throwaway rank
        # rather than writing temporary.txt into the project root.
        bundle = self._bundle()
        with tempfile.TemporaryDirectory() as root:
            context = sandbox.enter(root)
            text = student_scraper.write(bundle)
            self.assertTrue(os.path.exists(context.temp_student_file))
            for heading in ("SOURCE: INSTANT DATA SCRAPER", "SOURCE: GITHUB "
                            "PROFILE", "SOURCE: PORTFOLIO", "SOURCE: RESUME"):
                self.assertIn(heading, text)
            self.assertIn("Anilyadav9989", text)
            self.assertIn("FULL VISIBLE PROFILE TEXT", text)
            student_scraper.clear()
            self.assertFalse(os.path.exists(context.temp_student_file))
            self.assertFalse(os.path.exists(context.temp_student_json))


if __name__ == "__main__":
    unittest.main(verbosity=2)
