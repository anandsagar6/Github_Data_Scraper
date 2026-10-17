"""Profile-README extraction: raw Markdown, badges, and what they mean.

A GitHub profile README writes its tech stack as images:

    ![Java](https://img.shields.io/badge/java-...&logo=openjdk)

An <img> contributes nothing to innerText, so reading the rendered page
returns an EMPTY Tech Stack.  That is why this profile yielded only
"Java" (from a sentence) and "GitHub" (from a stats heading), while C,
HTML5 and CSS3 were lost entirely.

The fixture is a real capture of github.com/mythreyareddyannadi.
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
                       "fixtures", "mythreya_profile.json")

TECH_STACK = (
    "# 💻 Tech Stack:\n"
    "![C](https://img.shields.io/badge/c-%2300599C.svg?style=for-the-badge"
    "&logo=c&logoColor=white) "
    "![Java](https://img.shields.io/badge/java-%23ED8B00.svg?style="
    "for-the-badge&logo=openjdk&logoColor=white) "
    "![HTML5](https://img.shields.io/badge/html5-%23E34F26.svg?logo=html5) "
    "![CSS3](https://img.shields.io/badge/css3-%231572B6.svg?logo=css3)\n")


def fixture():
    with open(FIXTURE, encoding="utf-8") as handle:
        return json.load(handle)


def mythreya_record():
    data = fixture()
    fields = data["fields"]
    candidate = Candidate(
        username="mythreyareddyannadi",
        profile_url="https://github.com/mythreyareddyannadi",
        keyword="College Student Hyderabad",
        source_file="college-student-hyderabad.xlsx")
    bundle = SourceBundle(candidate=candidate)
    bundle.github_ok = True
    bundle.github_fields = fields
    bundle.github_text = data["text"]
    bundle.github_links = list(fields.get("links", []))
    return StudentMerger().build(bundle)


# ======================================================= badge parsing ====
class TestBadgeParsing(unittest.TestCase):
    def test_alt_text_is_read(self):
        names = [badge["alt"] for badge in student_scraper.extract_badges(TECH_STACK)]
        self.assertEqual(names, ["C", "Java", "HTML5", "CSS3"])

    def test_shields_label_fallback(self):
        self.assertEqual(
            student_scraper.shields_label("https://img.shields.io/badge/java-%23ED8B00.svg"),
            "java")
        self.assertEqual(
            student_scraper.shields_label("https://img.shields.io/badge/html5-%23E34F26"
                            ".svg?logo=html5"), "html5")
        self.assertEqual(student_scraper.shields_label("https://example.com/x.png"), "")

    def test_logo_slug_fallback(self):
        self.assertEqual(
            student_scraper.logo_name("https://img.shields.io/badge/java-x.svg?logo=openjdk"),
            "openjdk")

    def test_stats_images_are_decoration(self):
        markdown = ("![](https://github-readme-stats.vercel.app/api?user=x)\n"
                    "![](https://github-readme-streak-stats.herokuapp.com/)\n"
                    "[![](https://visitcount.itsvg.in/api?id=x)](https://v)\n")
        for badge in student_scraper.extract_badges(markdown):
            self.assertTrue(badge["decoration"], badge["url"])
        self.assertEqual(student_scraper.badge_names(markdown), [])

    def test_linked_badge_resolves_to_its_destination(self):
        markdown = ("[![LinkedIn](https://img.shields.io/badge/LinkedIn-x.svg)]"
                    "(https://www.linkedin.com/in/mythreyareddy/)")
        self.assertEqual(student_scraper.markdown_links(markdown),
                         [("LinkedIn",
                           "https://www.linkedin.com/in/mythreyareddy/")])


# ================================================ technology classifying ==
class TestBadgeSkills(unittest.TestCase):
    def test_tech_stack_badges_become_skills(self):
        result = student_scraper.extract_skills([(TECH_STACK, "readme_raw")])
        self.assertEqual(sorted(result["languages"]), ["C", "Java"])
        self.assertEqual(sorted(result["frameworks"]), ["CSS3", "HTML5"])

    def test_html_css_are_not_programming_languages(self):
        result = student_scraper.extract_skills([(TECH_STACK, "readme_raw")])
        self.assertNotIn("HTML5", result["languages"])
        self.assertNotIn("CSS3", result["languages"])

    def test_python_react_docker_classify_correctly(self):
        markdown = ("## Tech Stack\n"
                    "![Python](https://img.shields.io/badge/python-x.svg) "
                    "![React](https://img.shields.io/badge/react-x.svg) "
                    "![Docker](https://img.shields.io/badge/docker-x.svg)\n")
        result = student_scraper.extract_skills([(markdown, "readme_raw")])
        self.assertEqual(result["languages"], ["Python"])
        self.assertEqual(result["frameworks"], ["React"])
        self.assertEqual(result["tools"], ["Docker"])

    def test_single_letter_badges_count(self):
        for letter in ("C", "R"):
            markdown = ("## Tech Stack\n![%s](https://img.shields.io/badge/"
                        "%s-x.svg)\n" % (letter, letter.lower()))
            self.assertIn(letter,
                          student_scraper.extract_skills([(markdown,
                                             "readme_raw")])["languages"],
                          letter)

    def test_single_letters_in_prose_still_rejected(self):
        for text in ("I can see the result", "I can C the result",
                     "R is important", "the R&D team"):
            found = student_scraper.extract_skills([(text, "bio")])["languages"]
            self.assertNotIn("C", found, text)
            self.assertNotIn("R", found, text)

    def test_social_badges_are_never_skills(self):
        markdown = ("## 🌐 Socials:\n"
                    "[![Instagram](https://img.shields.io/badge/Instagram-x"
                    ".svg)](https://instagram.com/x) "
                    "[![LinkedIn](https://img.shields.io/badge/LinkedIn-x"
                    ".svg)](https://linkedin.com/in/x) "
                    "[![X](https://img.shields.io/badge/X-black.svg)]"
                    "(https://x.com/x)\n")
        result = student_scraper.extract_skills([(markdown, "readme_raw")])
        for bucket in ("languages", "frameworks", "tools", "primary"):
            self.assertEqual(result[bucket], [], bucket)

    def test_dsa_is_detected_from_prose(self):
        self.assertIn("DSA", student_scraper.concept_names(
            "I'm currently learning DSA using java."))
        self.assertIn("DSA", student_scraper.concept_names(
            "studying Data Structures and Algorithms"))

    def test_github_stats_heading_is_not_a_tool(self):
        markdown = ("# 💻 Tech Stack:\n![Java](https://img.shields.io/badge/"
                    "java-x.svg)\n"
                    "# 📊 GitHub Stats:\n"
                    "![](https://github-readme-stats.vercel.app/api?user=x)\n")
        result = student_scraper.extract_skills([(markdown, "readme_raw")])
        self.assertEqual(result["tools"], [])
        self.assertIn("Java", result["languages"])

    def test_canonical_mapping(self):
        for raw, expected in (("openjdk", "Java"), ("html5", "HTML5"),
                              ("css3", "CSS3"), ("c", "C"), ("dsa", "DSA"),
                              ("nodedotjs", "Node.js"),
                              ("Instagram", ""), ("X", ""), ("", "")):
            self.assertEqual(student_scraper.canonical_skill(raw), expected, raw)


# ====================================================== markdown sections =
class TestMarkdownSections(unittest.TestCase):
    def test_decorative_headings(self):
        markdown = ("# 💫 About Me:\nhello\n"
                    "## 🌐 Socials:\nlinks\n"
                    "# 💻 Tech Stack:\nbadges\n")
        found = student_scraper.parse_sections(markdown)
        self.assertEqual(found.get("about"), "hello")
        self.assertEqual(found.get("contact"), "links")
        self.assertEqual(found.get("skills"), "badges")

    def test_unknown_markdown_heading_ends_the_section(self):
        markdown = "# Tech Stack\nJava\n# GitHub Stats\nnoise\n"
        found = student_scraper.parse_sections(markdown)
        self.assertEqual(found.get("skills"), "Java")
        self.assertIn("noise", found.get("_other", ""))

    def test_unknown_bare_heading_ends_the_section(self):
        """The rendered page writes headings without any '#'."""
        rendered = "💻 Tech Stack:\nJava\n📊 GitHub Stats:\nnoise\n"
        found = student_scraper.parse_sections(rendered)
        self.assertNotIn("noise", found.get("skills", ""))

    def test_inline_heading_still_works(self):
        found = student_scraper.parse_sections("Skills: Python, SQL")
        self.assertEqual(found.get("skills"), "Python, SQL")

    def test_real_fixture_sections(self):
        raw = fixture()["fields"]["readme_raw"]
        found = student_scraper.parse_sections(raw)
        self.assertIn("about", found)
        self.assertIn("skills", found)
        self.assertIn("img.shields.io", found["skills"])


# ========================================================= college / place
class TestCollegeExtraction(unittest.TestCase):
    def test_lowercase_institution_after_an_enrolment_cue(self):
        self.assertEqual(
            student_scraper.extract_institution(
                "Student at malla reddy engineering college, hyderabad, "
                "telangana."),
            "Malla Reddy Engineering College")

    def test_capitalised_institution_still_works(self):
        self.assertEqual(
            student_scraper.extract_institution(
                "B.Tech student at Lakshmi Narain College of Technology"),
            "Lakshmi Narain College of Technology")

    def test_generic_phrases_invent_nothing(self):
        for text in ("student at a university", "I am a college student",
                     "an ordinary data science college student at a "
                     "university.", "winner of the university hackathon"):
            self.assertEqual(student_scraper.extract_institution(text), "", text)

    def test_education_level_without_a_degree(self):
        self.assertEqual(
            student_scraper.education_level(
                "", "Student at malla reddy engineering college"),
            "College Student")
        self.assertEqual(student_scraper.education_level("B.Tech", ""), "Undergraduate")

    def test_college_place_is_parsed_from_the_tail(self):
        parsed = student_scraper.parse_location("hyderabad, telangana")
        self.assertEqual(parsed.city, "Hyderabad")
        self.assertEqual(parsed.state, "Telangana")
        self.assertEqual(parsed.country, "India")

    def test_structured_location_with_trailing_stop(self):
        parsed = student_scraper.parse_location("Hyderabad, telangana, india.")
        self.assertEqual(parsed.as_tuple(),
                         ("Hyderabad", "Telangana", "India"))


# =========================================================== the profile ==
class TestMythreyaRegression(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.record = mythreya_record()

    def test_identity_and_location(self):
        self.assertEqual(self.record["Student Name"], "mythreyareddy annadi")
        self.assertEqual(self.record["GitHub Profile"],
                         "https://github.com/mythreyareddyannadi")
        self.assertEqual(self.record["Current City"], "Hyderabad")
        self.assertEqual(self.record["State"], "Telangana")
        self.assertEqual(self.record["Country"], "India")
        self.assertEqual(self.record["country"], "India")

    def test_college(self):
        self.assertEqual(self.record["College Name"],
                         "Malla Reddy Engineering College")
        self.assertEqual(self.record["College City"], "Hyderabad")
        self.assertEqual(self.record["College State"], "Telangana")
        self.assertEqual(self.record["Current Education Level"],
                         "College Student")
        self.assertEqual(self.record["Current Employment Status"], "Student")

    def test_skills_recovered_from_badges(self):
        self.assertEqual(self.record["Programming Languages"], "Java, C")
        self.assertEqual(self.record["Technologies / Frameworks"],
                         "CSS3, HTML5")
        self.assertEqual(self.record["Primary Skills"], "Java, DSA")
        for skill in ("C", "HTML5", "CSS3"):
            self.assertIn(skill, self.record["Secondary Skills"])

    def test_github_is_not_a_claimed_tool(self):
        self.assertEqual(self.record["Tools / Software Known"], "")
        self.assertNotIn("GitHub", self.record["Primary Skills"])

    def test_linkedin(self):
        self.assertEqual(self.record["LinkedIn Profile"],
                         "https://www.linkedin.com/in/mythreyareddy")

    def test_social_profiles_are_not_a_portfolio(self):
        self.assertEqual(self.record["Portfolio Website"], "")

    def test_nothing_is_invented(self):
        for column in ("Phone Number", "WhatsApp Number", "PIN Code",
                       "Current Course", "Current Specialization",
                       "Current Year / Semester", "Expected Graduation Year",
                       "UG Degree", "PG Degree", "Current Company",
                       "Current Job Role", "Resume URL", "Total Experience"):
            self.assertEqual(self.record[column], "", column)

    def test_schema_unchanged(self):
        self.assertEqual(list(self.record.keys()), C.FINAL_COLUMNS)
        self.assertEqual(len(self.record), 65)
        for alias, canonical in C.ALIAS_COLUMNS.items():
            self.assertEqual(self.record[alias], self.record[canonical])


if __name__ == "__main__":
    unittest.main(verbosity=2)
