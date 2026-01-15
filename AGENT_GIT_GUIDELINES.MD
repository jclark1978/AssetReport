# AI Agent Git Workflow Guidelines

These instructions define how you (the coding AI agent) should work with Git in this project.

## 1. Role and priorities

You are a coding assistant that **must**:

1. Help maintain a clean, understandable Git history.
2. Protect against data loss and leaked secrets.
3. Encourage good habits, not just “whatever works.”

If the user asks you to do something that breaks these rules, warn them, explain the impact, and suggest a safer alternative.

---

## 2. Branch workflow

### 2.1 Never develop on `main`

* Treat `main` as a **protected**, always-stable branch.
* Do **not** commit directly to `main` unless the user explicitly instructs you and understands the risk.
* Default assumption: all work happens on feature branches.

### 2.2 Feature branches

When starting new work, you should:

1. Confirm the base branch is up to date:

   * Suggest: `git checkout main` then `git pull origin main`.
2. Create a feature branch:

   * Suggest branch names in the format:

     * `feature/short-description`
     * `bugfix/short-description`
     * `chore/short-description`

Examples:

* `feature/add-login-validation`
* `bugfix/fix-timezone-bug`

If the user is already on a feature branch, keep using it for related work.

---

## 3. Commit best practices

### 3.1 Commit size and scope

When preparing commits, you must:

* Prefer **small, focused commits** that represent a single logical change.
* Avoid mixing unrelated changes in one commit (for example, refactor + new feature + formatting all in one).

If the staged changes are too broad, suggest splitting them into multiple commits.

### 3.2 Commit messages

You must propose commit messages that:

* Are short, clear, and in the **imperative mood**.
* Complete the sentence “This commit will…”

Examples of **good** messages:

* `Add validation to login form`
* `Refactor user repository for clarity`
* `Fix null reference error in auth service`
* `Update README with setup instructions`

Avoid:

* `fix`
* `changes`
* `misc updates`
* `final` or `final version`

### 3.3 Staging and committing flow

When the user asks you to “commit”:

1. Run or suggest `git status` and **summarize**:

   * What files changed.
   * What is staged and unstaged.

2. Propose a **commit plan**, for example:

   * “Let’s create one commit for the new API endpoint and a separate commit for test updates.”

3. Suggest explicit commands:

   * `git add <files>`
   * `git commit -m "<message>"`

4. Wait for user confirmation before assuming the commit is done.

---

## 4. Working with remotes

### 4.1 Pull before work

When starting a session or before creating a branch, you should:

* Suggest updating the base branch:

  * `git checkout main`
  * `git pull origin main`

### 4.2 Push behavior

When the user is done with a set of changes:

* Encourage pushing the branch to the remote:

  * `git push origin <branch-name>`

* If this is the first push for a branch, you may suggest:

  * `git push -u origin <branch-name>`

### 4.3 Force pushes

You must treat `git push --force` as **dangerous**:

* Do **not** suggest `--force` unless:

  * The user explicitly asks for it, **and**
  * You explain that it rewrites history and can affect others.

Prefer `git push --force-with-lease` when force-pushing is truly required and the user agrees.

---

## 5. .gitignore and sensitive data

### 5.1 Files that should not be committed

You must help ensure that the following **do not** get committed:

* Build artifacts and generated files (for example `dist/`, `build/`, `out/`).
* Dependency folders (for example `node_modules/`, `.venv/`).
* Editor and OS junk (`.DS_Store`, `Thumbs.db`, `.idea/`, `.vscode/` if appropriate).
* Local environment files (`.env`, `.env.local`, etc.).

If you see these files untracked in `git status`, you should:

1. Propose adding them to `.gitignore`.
2. Provide a specific diff suggestion for `.gitignore`.

### 5.2 Secrets and credentials

You must **never** encourage committing:

* API keys
* Passwords
* Tokens
* Private keys
* Cloud credentials

If such values appear in code or config:

* Warn the user clearly.
* Recommend:

  * Moving them to environment variables or secret management.
  * Rotating any secrets that were already committed.

If the user asks how to “hide” a secret that has already been committed, you should:

* Explain that simply editing the file is not enough.
* Recommend rotating the secret and, if necessary, using history-rewrite tools (for example `git filter-repo`) to scrub it, depending on the user’s comfort level.

---

## 6. Merge, conflicts, and history

### 6.1 Merging vs rebasing

Default, safe recommendation:

* Use **merge** for integrating feature branches into `main` unless the user requests rebasing.

If the user is comfortable with rebasing, you may:

* Suggest `git rebase main` on a feature branch to keep history linear.
* Always warn that rebasing rewrites history and is best done on local / not-yet-shared branches.

### 6.2 Conflict handling

When there is a merge or rebase conflict:

1. Ask the user to show the conflicted file(s) or `git status`.

2. Explain the conflict markers conceptually.

3. Help them decide what the final content should be by:

   * Comparing “current” vs “incoming” changes.
   * Proposing a combined version if appropriate.

4. After resolution, remind them to:

   * `git add <resolved-files>`
   * Complete the merge or rebase:

     * `git commit` for merges
     * `git rebase --continue` for rebases

---

## 7. Tags and releases

When the user reaches a significant milestone (for example “first working version”, “production release”), you should suggest:

* Creating an annotated tag:

  ```bash
  git tag -a vX.Y.Z -m "Description of this release"
  git push origin vX.Y.Z
  ```

If the project already uses semantic versioning, keep the same pattern.

---

## 8. Status and visibility

You should encourage frequent use of:

* `git status`
  To show what has changed and what is staged.

* `git diff` and `git diff --cached`
  To review changes before committing.

* `git log --oneline --graph --all`
  When the user seems confused about branches or history, suggest this to visualize the repo state.

Whenever the user asks you to “clean up” or “see where we are”, start by suggesting one or more of these commands and then interpret the output for them.

---

## 9. Safety and “do nots”

You must:

* Avoid destructive commands unless explicitly requested and understood, for example:

  * `git reset --hard`
  * `git clean -fd`
* Explain the consequences before suggesting them.

If the user asks for a “hard reset” or other dangerous operation:

1. Confirm what they are trying to achieve.
2. Offer a safer alternative if one exists (for example, `git stash`, new branch, revert commit).
3. Only then propose the destructive command, with a clear explanation.

---

## 10. When in doubt

If the user request conflicts with these rules or is ambiguous:

1. Default to the **safer, less destructive** option.
2. Ask for clarification or explicitly explain the trade-offs.
3. Keep `main` stable, history understandable, and secrets protected as top priorities.

---

## 11. Branch pruning policy

- Treat `main` (and any configured long-lived branches such as `develop`, `release/*`, `hotfix/*`) as permanent.
- Feature, bugfix, and chore branches (for example, `feature/...`, `bugfix/...`, `chore/...`) should be **deleted after they are merged into `main` and the changes are confirmed to be stable**.
- WIP or experiment branches (for example, `wip/...`, `sandbox/...`) may be kept temporarily, but should be periodically reviewed and either:
  - Cleaned up, merged, and then deleted, or
  - Deleted if abandoned.

When asked to “clean up branches” you should:

1. Ensure `main` is up to date (`git checkout main && git pull origin main`).
2. List branches merged into `main` (`git branch --merged main`).
3. Recommend deleting merged feature branches locally and remotely using:
   - `git branch -d <branch>` (local)
   - `git push origin --delete <branch>` (remote)

Never delete unmerged branches unless the user explicitly confirms they are no longer needed.

These rules are not suggestions; they are the default behavior you should follow when helping manage this repository with Git.

