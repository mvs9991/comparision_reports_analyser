"""Install the /recon-summary Claude Code skill.

The skill's source lives in claude_skill\\recon-summary\\ - a normal folder, so
it is committed and copied like any other file. Claude Code only loads skills
from a .claude\\skills folder, so this script copies it there and writes this
tool folder's location into the installed copy.

    python install_skill.py              install for your user: /recon-summary
                                         works in Claude Code opened in ANY folder
    python install_skill.py --project    install into this tool folder only
                                         (<tool folder>\\.claude\\skills)
    python install_skill.py --remove    uninstall (add --project for that copy)

Run it again after moving the tool folder or pulling a newer SKILL.md.
"""

import argparse
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_NAME = "recon-summary"
SOURCE = os.path.join(HERE, "claude_skill", SKILL_NAME)
PLACEHOLDER = "{{TOOL_FOLDER}}"


def main():
    ap = argparse.ArgumentParser(description="Install the /recon-summary Claude Code skill")
    ap.add_argument("--project", action="store_true",
                    help="install into this tool folder's .claude\\skills instead of your user folder")
    ap.add_argument("--remove", action="store_true", help="uninstall the skill")
    args = ap.parse_args()

    base = HERE if args.project else os.path.expanduser("~")
    target = os.path.join(base, ".claude", "skills", SKILL_NAME)

    if args.remove:
        if os.path.isdir(target):
            shutil.rmtree(target)
            print(f"Removed {target}")
        else:
            print(f"Nothing to remove - {target} does not exist")
        return

    skill_md = os.path.join(SOURCE, "SKILL.md")
    if not os.path.isfile(skill_md):
        sys.exit(f"ERROR: {skill_md} not found - run this from the tool folder that holds claude_skill\\")

    if os.path.isdir(target):
        shutil.rmtree(target)  # a clean copy, so a file removed from the source does not linger
    shutil.copytree(SOURCE, target)

    installed = os.path.join(target, "SKILL.md")
    with open(installed, encoding="utf-8") as fh:
        text = fh.read()
    with open(installed, "w", encoding="utf-8") as fh:
        fh.write(text.replace(PLACEHOLDER, HERE))

    print(f"Installed the /{SKILL_NAME} skill to:\n  {target}")
    print(f"It runs the tool from:\n  {HERE}")
    where = "Claude Code opened in this tool folder" if args.project else "Claude Code opened in any folder"
    print(f"Restart Claude Code, then in {where} type:  /{SKILL_NAME} <run-folder>")


if __name__ == "__main__":
    main()
