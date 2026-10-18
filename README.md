# CheckMate evidence enriched dataset

**Study 4 dataset repository:** https://github.com/chenxingyuzealken/CheckMate-Study4-Dataset

**Original CheckMate project repository:** https://github.com/bettersg/checkMate

Research data supplement for Study 4 on early information-disorder detection using CheckMate claims, target-directed emotion and stance analysis (TESA), related Reddit discussions, and retrieved webpages.

## Dataset

[Download the Excel workbook](Study4_reconstruction_dataset_TARGETS_VISIBLE_REVIEW.xlsx)

| Worksheet | Records | Description |
| --- | ---: | --- |
| Claims | 1,069 | Original CheckMate message ID, month, binary study label, data split, evidence counts |
| TESA | 21,286 | Target-directed emotion/stance annotations, entity type, target name or withholding flag |
| Reddit_Posts | 590 | Reddit post IDs, constructed URLs, timestamps and relevance metadata |
| Reddit_Comments | 17,416 | Reddit comment IDs, parent IDs and constructed URLs, with annotation linkage |
| Web_Evidence | 3,568 | Retrieved webpage URLs, ranks, dates and relevance scores |
| Queries | 3,568 | Query references; original search query text withheld unless approved |
| Codebook, Provenance, QA | — | Data dictionary, construction notes and known limitations |

**Join key:** `checkmate_message_id` is the *original* CheckMate message ID and links the data sheets. It is **not** an anonymised ID. The `tesa_ref` distinguishes individual target annotations; Reddit post and comment IDs link to their source nodes.

**TESA targets:** 21,083 extracted target labels are visible. A further 203 unusually long, URL-like, or potentially sensitive/source-text-like targets are withheld and flagged as `WITHHELD_REVIEW`. The raw withheld values are not distributed.

Temporal cutoffs are **0m, 1m, 15m, 20m, 60m, 90m, 24h, 48h, Full**. The `first_available_stage` indicator is estimated from source metadata; `Full` includes material with missing or late dates and does not mean it existed at claim submission.


## What is *not* included

The workbook **does not distribute the original CheckMate claim text**, exact CheckMate submission timestamps, Reddit body or author text, full webpage text, original unapproved search queries, the restricted modelling JSONL or private review/crosswalk material. Some links have their query strings removed to limit accidental publication of tokens or private data.

Source IDs and URLs can sometimes permit external re-collection **where resources remain publicly accessible and platform terms permit**. These are identifiers and pointers, not permission to republish third-party content. Retrieved webpages are not expert-verified fact-checks; the Reddit records describe related online discussion rather than direct WhatsApp forwarding traces.



## Important limitations and reuse notice

- The original CheckMate IDs and readable named-entity targets may create linkability or privacy risks; review handling of such data before onward use.
- Web evidence links have not been independently verified for current availability, safety, or source permissions; 3,568 link records are included.
- Some URL query strings or fragments were stripped and may affect reconstruction.
- **No licence has been granted in this repository.** No third-party source text or data are relicensed. Public accessibility is not a grant of reuse rights.
- The workbook's internal review/QA flags remain for transparency. They do not guarantee privacy clearance, verified links or source licensing.

Dataset prepared for the user-approved GitHub release on 18 October 2026. Please cite the associated Study 4 publication when available.
