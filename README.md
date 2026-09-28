# Outlook Inherit Category

A lightweight Outlook VBA macro that automatically keeps email conversations consistently color-coded — when a new message arrives, it looks through the rest of that thread and copies the most recently used category onto the new mail. No more manually re-tagging every reply in a long thread.

## Why this exists

Outlook doesn't tag replies with the same category as the rest of their conversation — every new message in a thread starts uncategorized, even if you've been carefully color-coding that thread for organization, priority, or client tracking. For anyone managing high email volume across ongoing threads (project coordination, vendor follow-ups, client conversations), that means constantly re-applying the same category by hand, every single reply.

This script automates that: it checks the conversation history across *all* folders and mailboxes (including shared mailboxes) the moment new mail lands, and silently inherits the right category — so your inbox stays organized without extra clicks.

## How it works

- Triggers on every new mail arrival (`Application_NewMailEx`)
- Resolves the item even if it landed in a non-default store (e.g. a shared mailbox)
- Uses Outlook's native Conversation object to scan the full thread, across folders
- Copies over the most recently applied category, skipping the new mail itself
- Fails silently on conflicts (e.g. a rule moving the item at the same time) instead of throwing errors

## 📁 File Structure
* `InheritCategory.cls` – the VBA logic

## ⚙️ Requirements
* Microsoft Outlook (desktop, Windows or Mac)
* Macros enabled in the Trust Center

## 🚀 How to Install

1. Open Outlook and press `Alt + F11` (Windows) or `Option + F11` (Mac) to open the VBA Editor.
2. In the Project pane, expand **Project1 → Microsoft Outlook Objects**.
3. Double-click **ThisOutlookSession**.
4. Paste the entire contents of `InheritCategory.cls` into that window.

   **Note:** the code must live inside `ThisOutlookSession` itself — it will not run if imported or left as a separate class module.

5. Restart Outlook to apply the changes.

## 🛡️ Disclaimer
Applies categories automatically and silently to incoming mail. Test in a non-production mailbox first and back up your VBA project before deploying. Provided as-is, without warranty of any kind.

## 📄 License
MIT
