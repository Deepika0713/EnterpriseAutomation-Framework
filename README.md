# Enterprise Growth Campaign & Lifecycle Automation Framework

A serverless, high-throughput marketing outreach and lifecycle automation framework built natively inside Google Workspace. This system abstracts backend script complexities into an intuitive spreadsheet dashboard, allowing non-technical teams to deploy cross-client marketing campaigns at scale without operational data loss.

## 🚀 Key Features & Architectural Highlights

* **Dynamic Data Schema Pipelines:** Powered by a custom, runtime regular-expression parsing engine `(/{{\s*(.+?)\s*}}/g)` that dynamically extracts user-segmentation parameters, maps payload metadata, and auto-generates structural database columns on the fly.
* **Idempotent State-Management:** Features strict row-by-row dispatch verification loops to monitor transmission states, gracefully catch null pointer exceptions, and completely eliminate duplicate message delivery.
* **High-Fidelity UI Engine:** Converts raw, complex dataset strings into fully responsive, semantic HTML layouts ensuring uniform UI/UX rendering across diverse email clients.
* **Low-Latency Optimization:** Engineered to run smoothly within cloud environment resource limitations, keeping script executions optimized for large enterprise data sets.

## 🛠️ Tech Stack & Ecosystem

* **Core Scripting:** JavaScript / Google Apps Script
* **Data & Analytics:** Python, Advanced Excel formulas
* **Database & Schemas:** SQL, JSON structured payloads
* **Frontend Components:** HTML5, CSS3 (Semantic markup)

## 📁 Repository Structure

├── src/

│   ├── Core.js          # Main execution loop and event triggers

│   ├── RegexEngine.js   # Dynamic parameter parsing and column building

│   └── StateManager.js  # Idempotent tracking and exception handling

├── templates/

│   └── Layout.html      # Responsive cross-client HTML canvas

├── tests/

│   └── MockData.json    # Sample schema arrays for pipeline verification


└── README.md
## ⚙️ How It Works: The Pipeline Under the Hood

The framework functions as a decoupled extract-transform-load (ETL) pipeline within the sheet UI:

1. **Ingestion & Validation:** The pipeline ingests user data arrays, applying real-time validation gates to clean missing values and structural breaks.
2. **Regex Evaluation:** The internal engine parses user-defined placeholders matching the `{{Variable}}` pattern to compute required database schema extensions.
3. **Execution & Logging:** The dispatch engine pushes variables into the HTML template canvas, fires the payload, and permanently stamps an immutable success token to the source data row.
## β Beta-version deploys
* **Email Automation _basic_
* https://docs.google.com/spreadsheets/d/1ZuXmW-eOgfh8iMQLDtXbf3I3d3Z9WDAwJZ4IMFMRSWw/edit?gid=1959928823#gid=1959928823
* **Auto-Computing Emails _basic_
* https://docs.google.com/spreadsheets/d/1Jt50RKYLX5Pf3nAWE8jG8h3z9LHFWdYWjyksgNyIWVU/edit?gid=0#gid=0
## 🛡️ Best Practices & Quality Control

* **Zero-Dependency Architecture:** Built completely utilizing native cloud runtime capabilities to minimize third-party API vulnerabilities and dependency rot.
* **Strict Idempotency:** The execution state is continually checked against a central transactional log before any external server network calls are allowed, guaranteeing system reliability.
