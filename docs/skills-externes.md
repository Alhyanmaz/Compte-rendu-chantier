# Skills externes installés dans le dépôt

Les sessions cloud repartent d'un clone du dépôt : seuls les skills commités dans `.claude/` y sont chargés
(les plugins déclarés dans `.claude/settings.json` ne le sont pas). Les deux outils ci-dessous sont donc copiés
dans le dépôt, à la version indiquée. Ils ne se mettent pas à jour tout seuls.

## Superpowers (Jesse Vincent, licence MIT)

- Source : https://github.com/obra/superpowers, v6.4.2, commit `8ca22db` du 25/09/2026.
- Installé : les 15 skills de `skills/`, copiés tels quels dans `.claude/skills/` (brainstorming, writing-plans,
  executing-plans, subagent-driven-development, test-driven-development, systematic-debugging,
  verification-before-completion, requesting-code-review, receiving-code-review, using-git-worktrees,
  finishing-a-development-branch, dispatching-parallel-agents, writing-skills, using-superpowers,
  diagnosing-superpowers).
- Non installé : le hook SessionStart du plugin, qui injecte `using-superpowers` au début de chaque session et
  impose de passer par les skills avant toute réponse. Sans lui, chaque skill se déclenche sur sa description.
- Les skills se citent entre eux sous la forme `superpowers:<nom>` ; dans ce dépôt, leur nom est `<nom>`.
- Télémétrie : seul le compagnon visuel optionnel de `brainstorming` charge un logo depuis le site de l'auteur.
  Variable `SUPERPOWERS_DISABLE_TELEMETRY=1` pour la couper.

## Hyperresearch (Jordan Gibbs, licence MIT)

- Source : https://github.com/jordan-gibbs/hyperresearch, paquet PyPI `hyperresearch` 0.12.0, commit `7790041`
  du 30/09/2026.
- Installé par `hyperresearch install .` : skill d'entrée `.claude/skills/hyperresearch/`, 18 skills d'étape
  `hyperresearch-*`, 16 agents `.claude/agents/hyperresearch-*.md`, bloc de documentation dans `CLAUDE.md`,
  configuration du vault dans `.hyperresearch/`.
- Utilisation : `/hyperresearch <question>`. Les notes et rapports sont écrits dans `research/` (à commiter pour
  les garder d'une session à l'autre). L'index SQLite `.hyperresearch/*.db` n'est pas versionné : la CLI le recrée.
- Fournisseur de récupération : `builtin` (HTTP simple) au lieu de `crawl4ai`, dont le navigateur headless ne
  correspond pas au Chromium préinstallé du conteneur. Les pages qui exigent JavaScript ne sont pas lues.
- Non installé : le hook PreToolUse qui rappelle de consulter le vault avant chaque WebSearch / WebFetch
  (`.hyperresearch/hook.js`, présent mais non branché).

### Prérequis dans l'environnement cloud

1. **CLI Python** : la commande `hyperresearch` n'existe pas au démarrage d'une session. Ajouter au script de
   configuration de l'environnement (menu de l'environnement dans la barre de titre de la session, puis Edit) :
   `pip install hyperresearch`. Le cache d'environnement la conserve ensuite d'une session à l'autre.
2. **Accès réseau** : la politique actuelle refuse les sites web courants (testé le 02/10/2026 : legifrance.gouv.fr,
   economie.gouv.fr, fr.wikipedia.org, github.com en HTTP 403). Une recherche Hyperresearch lit des dizaines de
   domaines imprévisibles : il faut l'accès réseau complet (Network access dans les réglages de l'environnement),
   une liste de domaines autorisés ne suffit pas. Voir
   https://code.claude.com/docs/en/cloud-environments#network-access.

## Mise à jour

Recopier `skills/` de Superpowers dans `.claude/skills/` ; pour Hyperresearch, `pip install -U hyperresearch` puis
`hyperresearch install .` (réécrit les skills, les agents et le bloc de `CLAUDE.md` ; supprimer ensuite le
`.claude/settings.json` qu'il crée si l'on ne veut pas du hook, et remettre `provider = "builtin"`).

Licences : `docs/licences/`.
