"""Unit tests for the extraction and normalisation rules.

Each block has positive AND negative cases: the negatives are what stop the
regexes from inventing data.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper                                           # noqa: E402



class TestEmail(unittest.TestCase):
    def test_positive(self):
        self.assertEqual(student_scraper.extract_emails("write to Anil.P9@Gmail.COM now"),
                         ["anil.p9@gmail.com"])
        self.assertEqual(student_scraper.extract_emails("a_b+tag@sub.domain.co.in"),
                         ["a_b+tag@sub.domain.co.in"])

    def test_negative(self):
        for text in ("example@example.com", "test@test.com",
                     "name@email.com", "bot@users.noreply.github.com",
                     "no-reply@github.com", "not an email @ all",
                     "user@localhost", "a@b.c"):
            self.assertEqual(student_scraper.extract_emails(text), [], text)

    def test_boundaries(self):
        # The local part is never truncated by a leading word character,
        # and trailing sentence punctuation is not swallowed.
        self.assertEqual(student_scraper.extract_emails("mail: foo.bar+x@sub.site.com."),
                         ["foo.bar+x@sub.site.com"])
        self.assertEqual(student_scraper.extract_emails("(reach me@site.org)"),
                         ["me@site.org"])
        # An address embedded in a filename is not a contact address.
        self.assertEqual(student_scraper.extract_emails("icon@2x.png"), [])


class TestPhone(unittest.TestCase):
    def test_formats(self):
        for text in ("Phone: 9876543210", "Mobile: 98765 43210",
                     "contact 98765-43210", "Phone +91 9876543210",
                     "call +91-98765-43210", "tel 91 9876543210"):
            found = student_scraper.extract_phones(text)
            self.assertTrue(found, text)
            self.assertEqual(found[0][0], "9876543210", text)

    def test_normalisation(self):
        self.assertEqual(student_scraper.normalize_phone("+91 98765 43210"), "9876543210")
        self.assertEqual(student_scraper.normalize_phone("09876543210"), "9876543210")
        self.assertEqual(student_scraper.normalize_phone("919876543210"), "9876543210")

    def test_rejects_non_phones(self):
        self.assertEqual(student_scraper.normalize_phone("1234567890"), "")   # starts 1
        self.assertEqual(student_scraper.normalize_phone("987654321"), "")    # 9 digits
        self.assertEqual(student_scraper.extract_phones("issue id 9876543210"), [])
        self.assertEqual(student_scraper.extract_phones("stars 9876543210"), [])
        # No context at all -> not accepted.
        self.assertEqual(student_scraper.extract_phones("9876543210 somewhere"), [])

    def test_whatsapp_only_when_labelled(self):
        self.assertEqual(student_scraper.extract_whatsapp("WhatsApp: 9123456789"),
                         "9123456789")
        self.assertEqual(student_scraper.extract_whatsapp("Mobile: 9123456789"), "")


class TestPin(unittest.TestCase):
    def test_positive(self):
        self.assertEqual(student_scraper.extract_pin("Address: Bhopal PIN 462001"),
                         "462001")
        self.assertEqual(student_scraper.extract_pin("Pincode 110001, India"), "110001")

    def test_negative(self):
        self.assertEqual(student_scraper.extract_pin("462001"), "")
        self.assertEqual(student_scraper.extract_pin("issue 462001"), "")
        self.assertEqual(student_scraper.extract_pin("PIN 062001"), "")   # leading zero


class TestUrls(unittest.TestCase):
    def test_linkedin(self):
        self.assertEqual(
            student_scraper.extract_linkedin("https://www.linkedin.com/in/foo-bar/?utm_a=1"),
            "https://www.linkedin.com/in/foo-bar")
        self.assertEqual(
            student_scraper.extract_linkedin("https://in.linkedin.com/in/foo"),
            "https://www.linkedin.com/in/foo")
        self.assertEqual(student_scraper.extract_linkedin("in/anil-palagani-a253b1254"),
                         "https://www.linkedin.com/in/anil-palagani-a253b1254")

    def test_linkedin_company_rejected(self):
        self.assertEqual(
            student_scraper.extract_linkedin("https://www.linkedin.com/company/acme"), "")

    def test_github_profile(self):
        import config as C
        self.assertEqual(
            student_scraper.extract_github_profile("https://github.com/octocat",
                                     C.RESERVED_PATHS),
            "https://github.com/octocat")
        self.assertEqual(
            student_scraper.extract_github_profile("https://github.com/search?q=x",
                                     C.RESERVED_PATHS), "")

    def test_portfolio_prefers_github_website(self):
        links = ["https://twitter.com/x", "https://me.dev"]
        self.assertEqual(student_scraper.extract_portfolio(links, "https://mysite.io"),
                         "https://mysite.io")
        self.assertEqual(student_scraper.extract_portfolio(links, ""), "https://me.dev")
        self.assertEqual(
            student_scraper.extract_portfolio(["https://github.com/x"], ""), "")

    def test_resume_urls(self):
        self.assertEqual(
            student_scraper.extract_resume_urls("my resume https://s.com/My_Resume.pdf"),
            ["https://s.com/My_Resume.pdf"])
        self.assertIn("https://s.com/cv",
                      student_scraper.extract_resume_urls("see https://s.com/cv"))
        self.assertEqual(student_scraper.extract_resume_urls("https://s.com/index.html"), [])


class TestYears(unittest.TestCase):
    def test_graduation(self):
        cases = {
            "B.Tech CSE 2023-2027": "2027",
            "Class of 2026": "2026",
            "Batch 2025": "2025",
            "Expected Graduation 2027": "2027",
            "Graduating 2026": "2026",
            "LNCT '26": "2026",
            "CSE student, graduating in 2028": "2028",
        }
        for text, expected in cases.items():
            self.assertEqual(student_scraper.extract_graduation_year(text), expected, text)

    def test_graduation_negative(self):
        self.assertEqual(student_scraper.extract_graduation_year("I scored 2026 points"), "")
        self.assertEqual(student_scraper.extract_graduation_year("version 2024 release"), "")
        self.assertEqual(student_scraper.extract_graduation_year("no year here"), "")

    def test_year_of_study(self):
        # Normalised to the ordinal form the output schema uses
        # ("3rd Year"), not the spelled-out word.
        cases = {
            "3rd year student": "3rd Year",
            "final year": "Final Year",
            "pre-final year": "Pre-Final Year",
            "first year student": "1st Year",
            "sem 5": "Semester 5",
            "5th semester": "Semester 5",
            "Semester 7": "Semester 7",
        }
        for text, expected in cases.items():
            self.assertEqual(student_scraper.extract_year_of_study(text), expected, text)

    def test_year_of_study_negative(self):
        self.assertEqual(student_scraper.extract_year_of_study("5 years of experience"), "")


class TestDegreeSpecialization(unittest.TestCase):
    def test_degrees(self):
        cases = {
            "B.Tech": "B.Tech", "BTech": "B.Tech", "B Tech": "B.Tech",
            "Bachelor of Technology": "B.Tech",
            "BE": "B.E.", "B.E.": "B.E.",
            "Bachelor of Engineering": "B.E.",
            "BCA": "BCA", "Bachelor of Computer Applications": "BCA",
            "MCA": "MCA", "Master of Computer Applications": "MCA",
            "B.Sc": "B.Sc", "BSc": "B.Sc", "Bachelor of Science": "B.Sc",
            "M.Tech": "M.Tech", "MTech": "M.Tech",
            "Master of Technology": "M.Tech",
            "MBA": "MBA", "Master of Business Administration": "MBA",
            "Diploma": "Diploma",
        }
        for text, expected in cases.items():
            self.assertEqual(student_scraper.normalize_degree(text), expected, text)

    def test_degree_negative(self):
        self.assertEqual(student_scraper.normalize_degree("be careful with me"), "")
        self.assertEqual(student_scraper.normalize_degree("nothing here"), "")

    def test_specializations(self):
        cases = {
            "CSE": "Computer Science Engineering",
            "Computer Science & Engineering": "Computer Science Engineering",
            "Computer Science": "Computer Science Engineering",
            "AIML": "Artificial Intelligence & Machine Learning",
            "AI/ML": "Artificial Intelligence & Machine Learning",
            "Data Science": "Data Science",
            "Cybersecurity": "Cyber Security",
            "ECE": "Electronics and Communication Engineering",
            "Mechanical Engineering": "Mechanical Engineering",
        }
        for text, expected in cases.items():
            self.assertEqual(student_scraper.normalize_specialization(text), expected, text)

    def test_ambiguous_needs_context(self):
        self.assertEqual(student_scraper.normalize_specialization("IT"), "")
        self.assertEqual(student_scraper.normalize_specialization("B.Tech IT"),
                         "Information Technology")

    def test_ambiguous_context_must_be_nearby(self):
        """A README mentioning AI is not a degree specialisation."""
        prose = ("Software Engineer - Distributed Systems - Backend "
                 "Engineering - AI | Exploring Agentic AI and AI-enabled "
                 "platforms across cloud-native student-friendly systems")
        self.assertEqual(student_scraper.normalize_specialization(prose), "")
        self.assertEqual(student_scraper.normalize_specialization("pursuing B.Tech in AI"),
                         "Artificial Intelligence")

    def test_ai_ml_not_matched_inside_repo_name(self):
        self.assertEqual(
            student_scraper.normalize_specialization("my repo AI-ML-TNS-Project"), "")


class TestExperience(unittest.TestCase):
    def test_durations(self):
        cases = {
            "Fresher": 0,
            "no experience": 0,
            "1 year of experience as a developer": 12,
            "1 yr 6 mo experience": 18,
            "18 months working at Infosys": 18,
            "2 years experience": 24,
            "2.5 years experience": 30,
            "2-3 years experience": 24,
            "2 to 3 years experience": 24,
        }
        for text, expected in cases.items():
            self.assertEqual(student_scraper.extract_experience(text), expected, text)

    def test_education_duration_is_not_experience(self):
        self.assertEqual(student_scraper.extract_experience("B.Tech is a 4 years degree"),
                         -1)
        self.assertEqual(student_scraper.extract_experience("a 3 years diploma course"), -1)

    def test_format(self):
        self.assertEqual(student_scraper.format_experience(0), "Fresher")
        self.assertEqual(student_scraper.format_experience(18), "1 year 6 months")
        self.assertEqual(student_scraper.format_experience(24), "2 years")
        self.assertEqual(student_scraper.format_experience(-1), "")


class TestCompanyRole(unittest.TestCase):
    def test_positive(self):
        self.assertEqual(
            student_scraper.extract_company_role("Lead Engineer @NatwestGroup | Ex-Amazon"),
            ("Lead Engineer", "NatwestGroup"))
        self.assertEqual(
            student_scraper.extract_company_role("Software Engineer at Infosys"),
            ("Software Engineer", "Infosys"))

    def test_lowercase_org_mention(self):
        """GitHub renders an org mention lower-cased."""
        self.assertEqual(
            student_scraper.extract_company_role("Lead Engineer @natwestgroup | Ex-Amazon"),
            ("Lead Engineer", "natwestgroup"))

    def test_at_word_still_needs_a_capitalised_company(self):
        self.assertEqual(student_scraper.extract_company_role("engineer at the office"),
                         ("", ""))

    def test_past_employer_rejected(self):
        self.assertEqual(student_scraper.extract_company_role("Ex-Amazon engineer"),
                         ("", ""))
        self.assertEqual(student_scraper.extract_company_role("formerly Developer at Wipro"),
                         ("", ""))

    def test_student_is_not_employment(self):
        self.assertEqual(student_scraper.extract_company_role("Student at LNCT"), ("", ""))


class TestLocation(unittest.TestCase):
    def test_indian(self):
        self.assertEqual(student_scraper.parse_location_tuple("Bhopal, Madhya Pradesh"),
                         ("Bhopal", "Madhya Pradesh", "India"))
        self.assertEqual(student_scraper.parse_location_tuple("Bangalore"),
                         ("Bengaluru", "Karnataka", "India"))
        self.assertEqual(student_scraper.parse_location_tuple("Bombay")[0], "Mumbai")
        self.assertEqual(student_scraper.parse_location_tuple("New Delhi")[2], "India")

    def test_country_not_invented(self):
        # A non-Indian place must not become India.
        self.assertEqual(student_scraper.parse_location_tuple("Trikala, Greece")[2], "Greece")
        # A recognised subdivision IS evidence: California means the US.
        # What must never happen is a default, so an unknown place stays
        # blank.
        self.assertEqual(student_scraper.parse_location_tuple("Santa Cruz, California")[2],
                         "United States")
        self.assertEqual(student_scraper.parse_location_tuple("Remote")[2], "")
        self.assertEqual(student_scraper.parse_location_tuple("Nowhereville")[2], "")


class TestSkills(unittest.TestCase):
    def test_finds(self):
        text = "Python, React, Node.js, Docker, C++ on Linux with Git"
        self.assertIn("Python",
                      student_scraper.find_skills(text, student_scraper.PROGRAMMING_LANGUAGES))
        self.assertIn("C++", student_scraper.find_skills(text, student_scraper.PROGRAMMING_LANGUAGES))
        self.assertIn("Node.js", student_scraper.find_skills(text, student_scraper.FRAMEWORKS))
        self.assertIn("Git", student_scraper.find_skills(text, student_scraper.TOOLS))

    def test_c_is_not_claimed_loosely(self):
        self.assertNotIn("C", student_scraper.find_skills("I can C the problem",
                                            student_scraper.PROGRAMMING_LANGUAGES))
        self.assertIn("C", student_scraper.find_skills("C and C++ developer",
                                         student_scraper.PROGRAMMING_LANGUAGES))

    def test_r_needs_a_programming_cue(self):
        self.assertNotIn("R", student_scraper.find_skills("Backend Engineering | R&D team",
                                            student_scraper.PROGRAMMING_LANGUAGES))
        self.assertIn("R", student_scraper.find_skills("Languages: Python, R, SQL",
                                         student_scraper.PROGRAMMING_LANGUAGES))

    def test_csharp_not_matched_as_c(self):
        found = student_scraper.find_skills("C# developer", student_scraper.PROGRAMMING_LANGUAGES)
        self.assertIn("C#", found)
        self.assertNotIn("C", found)


class TestCareer(unittest.TestCase):
    def test_interest(self):
        self.assertEqual(student_scraper.extract_career_interest("Aspiring DevOps Engineer"),
                         "DevOps Engineer")
        self.assertEqual(
            student_scraper.extract_career_interest("Interested in Cloud Computing"),
            "Cloud Computing")

    def test_preferred_role(self):
        self.assertEqual(student_scraper.extract_preferred_role("aspiring DevOps engineer"),
                         "DevOps Engineer")
        self.assertEqual(
            student_scraper.extract_preferred_role("looking for a Full Stack Developer role"),
            "Full Stack Developer")

    def test_employment_status(self):
        self.assertEqual(student_scraper.extract_employment_status("3rd year student"),
                         "Student")
        self.assertEqual(student_scraper.extract_employment_status("freelance developer"),
                         "Freelancer")
        self.assertEqual(student_scraper.extract_employment_status("", has_company=True),
                         "Employed")
        self.assertEqual(student_scraper.extract_employment_status("hello world"), "")


class TestCollege(unittest.TestCase):
    def test_positive(self):
        # An institution needs a cue that this is where the person studies;
        # a bare name with no context is not accepted, which is what stops
        # "a university" and certification-badge academies getting in.
        self.assertEqual(
            student_scraper.extract_college("B.Tech student at Lakshmi Narain College "
                              "of Technology"),
            "Lakshmi Narain College of Technology")
        self.assertIn("VIT", student_scraper.extract_college("Studying at VIT University"))
        self.assertEqual(
            student_scraper.extract_college("Lakshmi Narain College of Technology"), "")

    def test_negative(self):
        self.assertEqual(student_scraper.extract_college("just a bio with no school"), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
