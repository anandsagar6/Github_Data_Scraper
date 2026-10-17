"""Extraction accuracy: real profile regression plus false-positive guards.

The fixture is a genuine capture of https://github.com/LatiefDataVisionary,
stored so the suite runs offline.  Every "must stay blank" assertion below
corresponds to something the previous regex-only extractor got wrong.
"""

import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                   # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (Candidate, SourceBundle, StudentMerger)   # noqa: E402

FIXTURE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "fixtures", "latief_profile.json")


def latief_record():
    with open(FIXTURE, encoding="utf-8") as handle:
        data = json.load(handle)
    fields = data["fields"]
    candidate = Candidate(
        username="LatiefDataVisionary",
        profile_url="https://github.com/LatiefDataVisionary",
        keyword="College Student", source_file="college-student.xlsx",
        source_row_number=1, xlsx_name="Lathif Ramadhan",
        xlsx_bio=fields.get("bio", ""),
        raw_cells=["Lathif Ramadhan", "LatiefDataVisionary"])
    bundle = SourceBundle(candidate=candidate)
    bundle.github_ok = True
    bundle.github_fields = fields
    bundle.github_text = data["text"]
    bundle.github_links = list(fields.get("links", []))
    bundle.portfolio_url = student_scraper.extract_portfolio(bundle.github_links,
                                               fields.get("website", ""))
    return StudentMerger().build(bundle)


# ======================================================= 30. the regression
class TestLatiefRegression(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.record = latief_record()

    def test_identity(self):
        self.assertEqual(self.record["Student Name"], "Lathif Ramadhan")
        self.assertEqual(self.record["GitHub Profile"],
                         "https://github.com/LatiefDataVisionary")
        self.assertEqual(self.record["Email Address"],
                         "datasciencelatief@gmail.com")
        self.assertEqual(self.record["LinkedIn Profile"],
                         "https://www.linkedin.com/in/lathiframadhan")

    def test_location(self):
        self.assertEqual(self.record["Current City"], "Bantul")
        self.assertEqual(self.record["State"],
                         "Special Region of Yogyakarta")
        self.assertEqual(self.record["Country"], "Indonesia")
        self.assertEqual(self.record["country"], "Indonesia")

    def test_education(self):
        self.assertEqual(self.record["Current Education Level"],
                         "Undergraduate")
        self.assertEqual(self.record["Current Course"], "Data Science")
        self.assertEqual(self.record["Current Specialization"],
                         "Data Science")
        self.assertEqual(self.record["Current Year / Semester"], "3rd Year")

    def test_skills_are_real(self):
        languages = self.record["Programming Languages"]
        self.assertIn("Python", languages)
        self.assertIn("SQL", languages)
        frameworks = self.record["Technologies / Frameworks"]
        for expected in ("TensorFlow", "Pandas", "NumPy", "scikit-learn"):
            self.assertIn(expected, frameworks)
        self.assertIn("Git", self.record["Tools / Software Known"])

    def test_nothing_is_invented(self):
        """Every field the profile does not state must stay empty."""
        for column in ("Phone Number", "WhatsApp Number", "PIN Code",
                       "College Name", "University / Affiliation University",
                       "Expected Graduation Year", "Current Company",
                       "Current Job Role", "Resume URL", "Total Experience",
                       "College City", "College State"):
            self.assertEqual(self.record[column], "", column)

    def test_dicoding_is_not_a_portfolio(self):
        self.assertEqual(self.record["Portfolio Website"], "")
        self.assertEqual(self.record["portfolio"], "")

    def test_a_university_is_not_a_college_name(self):
        """The bio says "at a university" -- that is not an institution."""
        self.assertEqual(self.record["College Name"], "")
        self.assertNotIn("ordinary", self.record["College Name"])

    def test_employment_status_is_student(self):
        self.assertEqual(self.record["Current Employment Status"], "Student")

    def test_schema_unchanged(self):
        self.assertEqual(list(self.record.keys()), C.FINAL_COLUMNS)
        self.assertEqual(len(self.record), 65)
        for alias, canonical in C.ALIAS_COLUMNS.items():
            self.assertEqual(self.record[alias], self.record[canonical],
                             alias)


# ============================================ 31. false-positive regressions
class TestFalsePositives(unittest.TestCase):
    def test_1_ai_enthusiast_is_not_a_specialization(self):
        self.assertEqual(
            student_scraper.normalize_specialization("AI enthusiast building AI tools"),
            "")
        self.assertEqual(
            student_scraper.normalize_specialization("pursuing B.Tech in AI"),
            "Artificial Intelligence")

    def test_2_negated_skill_is_not_claimed(self):
        found = student_scraper.extract_skills(
            [("I use Python but no longer Java", "bio")])["languages"]
        self.assertIn("Python", found)
        self.assertNotIn("Java", found)

    def test_3_letter_r_in_prose(self):
        for text in ("Backend Engineering | R&D team", "Rated R for content",
                     "R is the 18th letter"):
            self.assertNotIn(
                "R", student_scraper.extract_skills([(text, "bio")])["languages"], text)
        self.assertIn("R", student_scraper.extract_skills(
            [("Skills\nPython, R, SQL", "readme")])["languages"])

    def test_4_letter_c_in_prose(self):
        self.assertNotIn("C", student_scraper.extract_skills(
            [("I can C the problem", "bio")])["languages"])
        self.assertIn("C", student_scraper.extract_skills(
            [("C and C++ developer", "bio")])["languages"])

    def test_5_university_hackathon_is_not_a_university(self):
        result = student_scraper.extract_education(
            [("winner of the university hackathon 2024", 0.9)])
        self.assertEqual(result["university"], "")
        self.assertEqual(result["college"], "")

    def test_6_student_at_a_university_invents_nothing(self):
        for text in ("student at a university", "I study at a college",
                     "an ordinary data science college student at a "
                     "university."):
            result = student_scraper.extract_education([(text, 0.9)])
            self.assertEqual(result["college"], "", text)
            self.assertEqual(result["university"], "", text)

    def test_7_lowercase_org_mention(self):
        self.assertEqual(
            student_scraper.extract_company_role("Lead Engineer @natwestgroup | Ex-Amazon"),
            ("Lead Engineer", "natwestgroup"))

    def test_8_student_at_lnct_is_not_employment(self):
        self.assertEqual(student_scraper.extract_company_role("Student at LNCT"), ("", ""))
        self.assertEqual(
            student_scraper.extract_company_role("3rd year student at LNCT"), ("", ""))
        # ...but it IS a college.
        self.assertIn("LNCT",
                      student_scraper.extract_education([("3rd year student at LNCT "
                                             "Bhopal", 0.9)])["college"])

    def test_9_btech_four_years_is_not_experience(self):
        self.assertEqual(student_scraper.extract_experience("B.Tech 4 years"), -1)
        self.assertEqual(
            student_scraper.format_experience(student_scraper.extract_experience("B.Tech 4 years")), "")

    def test_10_internship_does_not_become_total_experience(self):
        resume = ("EXPERIENCE\nSoftware Intern at Infosys, 6 months\n"
                  "2 years experience building web applications")
        self.assertEqual(student_scraper.extract_experience(resume), 24)
        internships = student_scraper.extract_internships(resume)
        self.assertTrue(any("Intern" in line for line in internships))

    def test_11_open_to_work_is_not_unemployed(self):
        status = student_scraper.extract_employment_status("Open to Work")
        self.assertNotEqual(status, "Unemployed")
        self.assertEqual(status, "Job Seeker")
        # An employed person may also say it.
        self.assertEqual(
            student_scraper.extract_employment_status("Open to Work", has_company=True),
            "Employed")

    def test_12_dicoding_is_not_a_portfolio(self):
        self.assertEqual(student_scraper.extract_portfolio(
            ["https://www.dicoding.com/users/x/academies"], ""), "")
        for url in ("https://leetcode.com/u/x", "https://codechef.com/users/x",
                    "https://www.credly.com/users/x"):
            self.assertEqual(student_scraper.extract_portfolio([url], ""), "", url)
        self.assertEqual(student_scraper.extract_portfolio(["https://lathif.dev"], ""),
                         "https://lathif.dev")


# ==================================================== targeted field tests
class TestLocationAccuracy(unittest.TestCase):
    def test_global_locations(self):
        cases = {
            "Bantul, Special Region of Yogyakarta":
                ("Bantul", "Special Region of Yogyakarta", "Indonesia"),
            "Bhopal, Madhya Pradesh": ("Bhopal", "Madhya Pradesh", "India"),
            "Bangalore": ("Bengaluru", "Karnataka", "India"),
            "Bombay": ("Mumbai", "Maharashtra", "India"),
            "Gurgaon": ("Gurugram", "Haryana", "India"),
            "Toronto, Canada": ("Toronto", "Ontario", "Canada"),
            # Bare "London" is ambiguous (London, Ontario), so the country
            # is deliberately left blank until context settles it.
            "London": ("London", "", ""),
            "London, UK": ("London", "", "United Kingdom"),
            "Jakarta, Indonesia": ("Jakarta", "Jakarta", "Indonesia"),
        }
        for raw, expected in cases.items():
            self.assertEqual(student_scraper.parse_location(raw).as_tuple(), expected, raw)

    def test_country_is_never_guessed(self):
        for raw in ("Remote", "Earth", "somewhere over the rainbow", ""):
            self.assertEqual(student_scraper.parse_location(raw).country, "", raw)
        # A non-Indian place must never become India.
        self.assertEqual(student_scraper.parse_location("Trikala, Greece").country,
                         "Greece")

    def test_priority_fills_gaps_only(self):
        best = student_scraper.best_location([("Bantul, Special Region of Yogyakarta", 1.0),
                                ("Mumbai, Maharashtra", 0.8)])
        self.assertEqual(best.city, "Bantul")
        self.assertEqual(best.country, "Indonesia")


class TestEducationAccuracy(unittest.TestCase):
    def test_degrees(self):
        for text, expected in {
                "B.Tech": "B.Tech", "BTech": "B.Tech",
                "Bachelor of Technology": "B.Tech", "BE": "B.E.",
                "Bachelor of Engineering": "B.E.", "BCA": "BCA",
                "MCA": "MCA", "BSc": "B.Sc", "MTech": "M.Tech",
                "MBA": "MBA", "Polytechnic": "Diploma"}.items():
            self.assertEqual(student_scraper.normalize_degree(text), expected, text)
        self.assertEqual(student_scraper.normalize_degree("be careful with me"), "")

    def test_specializations(self):
        for text, expected in {
                "CSE branch": "Computer Science Engineering",
                "B.Tech IT": "Information Technology",
                "AI/ML branch":
                    "Artificial Intelligence & Machine Learning",
                "Data Science": "Data Science",
                "pursuing ECE": "Electronics and Communication Engineering",
                }.items():
            self.assertEqual(student_scraper.normalize_specialization(text), expected,
                             text)

    def test_year_of_study(self):
        for text, expected in {
                "3rd year student": "3rd Year",
                "final year": "Final Year",
                "pre-final year": "Pre-Final Year",
                "sem 5": "Semester 5",
                "5th semester": "Semester 5",
                "3rd Data Science honours student": "3rd Year"}.items():
            self.assertEqual(student_scraper.extract_year_of_study(text), expected, text)

    def test_graduation_year(self):
        for text, expected in {
                "B.Tech CSE 2023-2027": "2027",
                "Class of 2026": "2026", "Batch 2025": "2025",
                "Expected Graduation 2027": "2027",
                "LNCT '26": "2026"}.items():
            self.assertEqual(student_scraper.extract_graduation_year(text), expected, text)

    def test_copyright_year_is_not_graduation(self):
        for text in ("Copyright 2024 all rights reserved",
                     "version 2024 release notes", "© 2025 My Site"):
            self.assertEqual(student_scraper.extract_graduation_year(text), "", text)

    def test_real_institution_names_survive(self):
        for text, expected in {
                "B.Tech student at Lakshmi Narain College of Technology":
                    "Lakshmi Narain College of Technology",
                "studying at VIT University": "VIT University",
                }.items():
            self.assertEqual(student_scraper.extract_institution(text), expected, text)

    def test_education_level(self):
        self.assertEqual(student_scraper.education_level("B.Tech", ""), "Undergraduate")
        self.assertEqual(student_scraper.education_level("MCA", ""), "Postgraduate")
        self.assertEqual(student_scraper.education_level("", "college student"),
                         "Undergraduate")
        self.assertEqual(student_scraper.education_level("", "high school student"),
                         "School Student")
        self.assertEqual(student_scraper.education_level("", "nothing relevant"), "")


class TestSectionParsing(unittest.TestCase):
    def test_headings_are_recognised(self):
        text = ("About Me\nI build things.\n"
                "Technical Skills\nPython, Django\n"
                "Education\nB.Tech CSE\n"
                "Experience\nIntern at ACME")
        found = student_scraper.parse_sections(text)
        self.assertIn("about", found)
        self.assertIn("skills", found)
        self.assertIn("education", found)
        self.assertIn("experience", found)
        self.assertIn("Python", found["skills"])
        self.assertIn("B.Tech", found["education"])

    def test_inline_heading(self):
        found = student_scraper.parse_sections("Skills: Python, SQL\nother line")
        self.assertIn("skills", found)
        self.assertIn("Python", found["skills"])

    def test_real_readme_sections(self):
        with open(FIXTURE, encoding="utf-8") as handle:
            readme = json.load(handle)["fields"]["readme"]
        found = student_scraper.parse_sections(readme)
        self.assertIn("skills", found)
        self.assertIn("Python", found["skills"])


class TestSkillsWeighting(unittest.TestCase):
    def test_skills_section_outranks_prose(self):
        scores = student_scraper.collect_skills([
            ("Skills\nPython, Django", "resume"),
            ("I once saw a Rust talk", "profile")])
        self.assertGreater(scores.get("Python", 0), scores.get("Rust", 0))

    def test_primary_reflects_evidence_not_order(self):
        result = student_scraper.extract_skills([
            ("Skills\nDocker, Kubernetes", "resume"),
            ("Skills\nDocker", "portfolio"),
            ("I mentioned Redis once", "profile")])
        self.assertIn("Docker", result["primary"])
        self.assertNotIn("Redis", result["primary"])

    def test_classification_buckets(self):
        result = student_scraper.extract_skills(
            [("Skills\nPython, React, Docker, Git", "resume")])
        self.assertIn("Python", result["languages"])
        self.assertIn("React", result["frameworks"])
        self.assertIn("Git", result["tools"])


class TestEmploymentAccuracy(unittest.TestCase):
    def test_company_and_role(self):
        cases = {
            "Software Engineer @ Microsoft": ("Software Engineer",
                                              "Microsoft"),
            "Developer at TCS": ("Developer", "TCS"),
            "Working as DevOps Engineer at ABC": ("DevOps Engineer", "ABC"),
            "Frontend Developer - XYZ": ("Frontend Developer", "XYZ"),
        }
        for text, expected in cases.items():
            self.assertEqual(student_scraper.extract_company_role(text), expected, text)

    def test_education_blocks_company(self):
        for text in ("Student at ABC University", "B.Tech at IIT Delhi",
                     "pursuing MCA at XYZ College"):
            self.assertEqual(student_scraper.extract_company_role(text), ("", ""), text)

    def test_career_interest_and_preferred_role(self):
        self.assertEqual(student_scraper.extract_career_interest("Aspiring DevOps Engineer"),
                         "DevOps Engineer")
        self.assertEqual(student_scraper.extract_preferred_role("aspiring DevOps engineer"),
                         "DevOps Engineer")
        self.assertEqual(
            student_scraper.extract_career_interest("Interested in Cloud Computing"),
            "Cloud Computing")

    def test_preferred_industry_needs_explicit_evidence(self):
        self.assertEqual(student_scraper.extract_preferred_industry("Cloud Computing"),
                         "Cloud Computing")
        # A language alone is never an industry.
        self.assertEqual(student_scraper.extract_preferred_industry("Python"), "")
        self.assertEqual(student_scraper.extract_preferred_industry("Docker React"), "")

    def test_experience_formats(self):
        for text, months in {
                "Fresher": 0, "no experience": 0,
                "1 year of experience as a developer": 12,
                "1 yr 6 mo experience": 18,
                "18 months working at Infosys": 18,
                "2.5 years experience": 30,
                "2-3 years experience": 24,
                "2 to 3 years experience": 24}.items():
            self.assertEqual(student_scraper.extract_experience(text), months, text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
