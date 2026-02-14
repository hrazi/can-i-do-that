import { useState, useEffect, useCallback } from "react";
import {
  useMsal,
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { loginRequest } from "./authConfig";

// ── Types ──

type VerdictType = "yes" | "caution" | "no" | "depends";

interface Source {
  text: string;
  url: string;
}

interface AIResponse {
  verdict: string;
  verdictType: VerdictType;
  explanation: string;
  considerations: string[];
  nextSteps: string[];
  sources: Source[];
}

// ── SVG Components ──

function GlassesIcon({
  size = 60,
  className = "",
}: {
  size?: number;
  className?: string;
}) {
  return (
    <svg
      width={size}
      height={size * 0.45}
      viewBox="0 0 120 50"
      fill="none"
      className={className}
      aria-hidden="true"
    >
      <circle cx="32" cy="26" r="19" stroke="currentColor" strokeWidth="3.5" />
      <circle cx="88" cy="26" r="19" stroke="currentColor" strokeWidth="3.5" />
      <path
        d="M51 24 Q60 16 69 24"
        stroke="currentColor"
        strokeWidth="3.5"
        fill="none"
        strokeLinecap="round"
      />
      <path
        d="M13 26 L2 22"
        stroke="currentColor"
        strokeWidth="3.5"
        strokeLinecap="round"
      />
      <path
        d="M107 26 L118 22"
        stroke="currentColor"
        strokeWidth="3.5"
        strokeLinecap="round"
      />
      <circle cx="25" cy="20" r="4" fill="currentColor" opacity="0.1" />
      <circle cx="81" cy="20" r="4" fill="currentColor" opacity="0.1" />
    </svg>
  );
}

function MicrosoftLogo() {
  return (
    <svg
      className="ms-logo-svg"
      viewBox="0 0 21 21"
      xmlns="http://www.w3.org/2000/svg"
    >
      <rect x="1" y="1" width="9" height="9" fill="#f25022" />
      <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
      <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
      <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
    </svg>
  );
}

// ── Example Questions ──

const EXAMPLE_QUESTIONS = [
  "Can I use Claude Code for internal development?",
  "Can I integrate Anthropic Claude into a Microsoft product?",
  "Can I generate synthetic data using Anthropic models?",
  "Can I use Anthropic models with customer data?",
  "Can I fine-tune a model using Claude outputs?",
  "Can I deploy Anthropic models in a government cloud?",
  "Can I use Claude from our China office?",
  "Can I paste OpenAI code into Claude?",
  "Can I use the Anthropic logo in our product?",
  "Can I use Claude with internal Microsoft documents?",
  "Can I add Claude to an M365 enterprise service?",
  "Can I use Anthropic models for EU customers?",
];

// ── Source URLs ──

const ANTHROPIC_GUIDANCE_URL =
  "https://docs.opensource.microsoft.com/legal/cela-guidance/anthropic-engineering-guidance/anthropic-engineering-guidance/";
const MS_RESPONSIBLE_AI_URL =
  "https://learn.microsoft.com/en-us/compliance/assurance/assurance-artificial-intelligence";
const MS_DATA_CLASSIFICATION_URL =
  "https://learn.microsoft.com/en-us/compliance/assurance/assurance-create-data-classification-framework";
const MS_DATA_LABELS_URL =
  "https://learn.microsoft.com/en-us/compliance/assurance/assurance-data-classification-and-labels";
const MS_PURVIEW_INFO_PROTECTION_URL =
  "https://learn.microsoft.com/en-us/purview/information-protection-solution";
const MS_GDPR_URL =
  "https://learn.microsoft.com/en-us/compliance/regulatory/gdpr";
const MS_AI_GOVERNANCE_URL =
  "https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/scenarios/ai/govern";
const MS_AI_REGULATORY_URL =
  "https://learn.microsoft.com/en-us/security/security-for-ai/govern";
const MS_PURVIEW_AI_URL =
  "https://learn.microsoft.com/en-us/purview/ai-microsoft-purview";

// ── Mock AI Evaluation ──

function evaluateScenario(question: string): Promise<AIResponse> {
  return new Promise((resolve) => {
    const q = question.toLowerCase();

    // Simulate network delay
    const delay = 1800 + Math.random() * 1200;

    setTimeout(() => {
      // ── Fine-tuning / distillation / training competing models (Section A) ──
      if (
        q.includes("fine-tun") ||
        q.includes("finetun") ||
        q.includes("distill") ||
        (q.includes("train") &&
          (q.includes("competing") ||
            q.includes("model") ||
            q.includes("our own")))
      ) {
        resolve({
          verdict: "Not allowed.",
          verdictType: "no",
          explanation:
            "Per Section A of the Engineering Guidance for Using Anthropic Models, using Anthropic model outputs to fine-tune, distill, or train competing models is expressly prohibited. This includes using Claude outputs as training data for any Microsoft or third-party model.",
          considerations: [
            "This restriction applies regardless of the data type or internal purpose",
            "Generating synthetic training data for competing models also falls under this prohibition",
            "Anthropic's Terms of Service explicitly forbid this use case",
          ],
          nextSteps: [
            "If you need to train a custom model, use Azure OpenAI or other approved first-party services",
            "Contact CELA if you have questions about what constitutes a 'competing model'",
            "Review the full Section A restrictions in the Engineering Guidance document",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section A: Prohibited Uses", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Govern AI — Cloud Adoption Framework", url: MS_AI_GOVERNANCE_URL },
          ],
        });
      }

      // ── Use from China / Hong Kong (Section A) ──
      else if (
        q.includes("china") ||
        q.includes("hong kong") ||
        q.includes("unsupported countr")
      ) {
        resolve({
          verdict: "Not allowed.",
          verdictType: "no",
          explanation:
            "Per Section A of the Engineering Guidance for Using Anthropic Models, Anthropic does not support usage from China, Hong Kong, or other unsupported countries. Microsoft employees located in or accessing services from these regions cannot use Anthropic models.",
          considerations: [
            "This is an Anthropic platform-level restriction, not a Microsoft policy choice",
            "VPN usage does not change the compliance status — location of the user matters",
            "This applies to all Anthropic model access including API and Claude Code",
          ],
          nextSteps: [
            "Use Azure OpenAI or other approved alternatives for work in these regions",
            "Check the Anthropic supported countries list for the latest restrictions",
            "Contact CELA if you need guidance on region-specific AI tool availability",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section A: Unsupported Countries", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Govern AI apps and data for regulatory compliance", url: MS_AI_REGULATORY_URL },
          ],
        });
      }

      // ── Government / sovereign cloud (Section D.6) ──
      else if (
        q.includes("government") ||
        q.includes("gov cloud") ||
        q.includes("fedramp") ||
        q.includes("sovereign") ||
        q.includes("gcc") ||
        q.includes("dod")
      ) {
        resolve({
          verdict: "Not allowed.",
          verdictType: "no",
          explanation:
            "Per Section D.6 of the Engineering Guidance for Using Anthropic Models, Anthropic models cannot be deployed in Government or Sovereign cloud environments. Anthropic does not hold FedRAMP authorization, making deployment in these environments non-compliant.",
          considerations: [
            "This applies to all US Government cloud tiers (GCC, GCC High, DoD)",
            "Sovereign cloud environments worldwide are also excluded",
            "No timeline has been provided for FedRAMP authorization from Anthropic",
          ],
          nextSteps: [
            "Use Azure OpenAI for Government cloud workloads — it holds FedRAMP authorization",
            "Contact the Azure Government team for approved AI alternatives",
            "Review Section D.6 of the guidance for the full list of excluded environments",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section D.6: Government & Sovereign Cloud", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Govern AI apps and data for regulatory compliance", url: MS_AI_REGULATORY_URL },
          ],
        });
      }

      // ── OpenAI / proprietary code into Claude (Section B) ──
      else if (
        (q.includes("openai") && q.includes("code")) ||
        (q.includes("proprietary") && q.includes("code")) ||
        (q.includes("third-party") && q.includes("code")) ||
        (q.includes("paste") && q.includes("code"))
      ) {
        resolve({
          verdict: "Not allowed.",
          verdictType: "no",
          explanation:
            "Per Section B of the Engineering Guidance for Using Anthropic Models, feeding OpenAI proprietary code or other third-party proprietary code into Claude or Anthropic models is prohibited. This protects Microsoft from IP contamination and contractual violations.",
          considerations: [
            "This applies to code from OpenAI, Google, Meta, and other third-party proprietary sources",
            "Using Claude Code on repositories containing such code is also restricted",
            "CELA review is required before using any third-party proprietary code with Anthropic models",
          ],
          nextSteps: [
            "Ensure your codebase does not contain third-party proprietary code before using Claude Code",
            "Submit to the Third Party AI Models CELA Review site if you need an exception",
            "Contact CELA for guidance on what constitutes 'proprietary code' in your context",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section B: Proprietary Code Restrictions", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Microsoft Purview data security for generative AI apps", url: MS_PURVIEW_AI_URL },
          ],
        });
      }

      // ── Claude Code for internal development (Section B) ──
      else if (
        q.includes("claude code") ||
        (q.includes("claude") && q.includes("develop")) ||
        (q.includes("claude") && q.includes("coding")) ||
        (q.includes("anthropic") && q.includes("develop") && !q.includes("product"))
      ) {
        resolve({
          verdict: "Yes, you can do that!",
          verdictType: "yes",
          explanation:
            "Per Section B of the Engineering Guidance for Using Anthropic Models, Claude Code is approved for internal development use. Follow the 1ES (One Engineering System) guidance for setup. No CELA review is needed for standard internal development use.",
          considerations: [
            "Follow the 1ES guidance for Claude Code configuration and setup",
            "Do not feed OpenAI or other third-party proprietary code into Claude Code",
            "Use Zero Data Retention (ZDR) if working with sensitive or confidential data",
            "The standard procurement process should be followed for licensing",
          ],
          nextSteps: [
            "Follow the 1ES onboarding guide for Claude Code setup",
            "Ensure your team has completed the Anthropic procurement process",
            "Contact Steven Stanford (sstanford@microsoft.com) for ZDR setup if handling sensitive data",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section B: Internal Development & Claude Code", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Govern AI — Cloud Adoption Framework", url: MS_AI_GOVERNANCE_URL },
          ],
        });
      }

      // ── Internal testing / evaluation (Section B) ──
      else if (
        (q.includes("test") || q.includes("eval") || q.includes("benchmark")) &&
        (q.includes("internal") ||
          q.includes("anthropic") ||
          q.includes("claude"))
      ) {
        resolve({
          verdict: "Yes, you can do that!",
          verdictType: "yes",
          explanation:
            "Per Section B of the Engineering Guidance for Using Anthropic Models, internal testing and evaluation of Anthropic models is permitted. This includes benchmarking, quality assessments, and proof-of-concept work, provided you follow the procurement process.",
          considerations: [
            "Complete the Anthropic procurement process before beginning evaluations",
            "Use ZDR (Zero Data Retention) if testing with sensitive or confidential data",
            "Evaluation results may be shared internally but check before publishing externally",
          ],
          nextSteps: [
            "Ensure your team has access through the approved procurement channel",
            "Contact Steven Stanford (sstanford@microsoft.com) for ZDR setup if needed",
            "Document evaluation methodology and results for compliance records",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section B: Internal Testing & Evaluation", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Microsoft Responsible AI overview", url: MS_RESPONSIBLE_AI_URL },
          ],
        });
      }

      // ── Internal Microsoft data, non-customer (Section B) ──
      else if (
        q.includes("internal data") ||
        q.includes("internal microsoft") ||
        (q.includes("internal") &&
          !q.includes("customer") &&
          !q.includes("product") &&
          (q.includes("data") || q.includes("document") || q.includes("info")))
      ) {
        resolve({
          verdict: "Yes, with conditions.",
          verdictType: "yes",
          explanation:
            "Per Section B of the Engineering Guidance for Using Anthropic Models, using Anthropic models with internal Microsoft data (non-customer) is permitted. However, Zero Data Retention (ZDR) is required when working with sensitive or confidential internal data.",
          considerations: [
            "ZDR must be enabled for any sensitive, confidential, or highly confidential data",
            "Standard (non-sensitive) internal data can be used without ZDR",
            "Do not include customer data — that requires additional approvals (see Section C.2)",
            "Ensure data classification is correct before processing",
          ],
          nextSteps: [
            "Classify your data per Microsoft's data classification framework",
            "Contact Steven Stanford (sstanford@microsoft.com) to set up ZDR if handling sensitive data",
            "Complete the Anthropic procurement process if your team hasn't already",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section B: Internal Data & ZDR Requirements", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Data classification framework", url: MS_DATA_CLASSIFICATION_URL },
            { text: "Data classification & sensitivity labels", url: MS_DATA_LABELS_URL },
          ],
        });
      }

      // ── Product integration — enterprise or consumer (Section C) ──
      else if (
        q.includes("product") ||
        q.includes("integrat") ||
        q.includes("ship") ||
        q.includes("embed") ||
        (q.includes("build") &&
          (q.includes("feature") || q.includes("app") || q.includes("service")))
      ) {
        resolve({
          verdict: "Hold on — proceed with caution!",
          verdictType: "caution",
          explanation:
            "Per Section C of the Engineering Guidance for Using Anthropic Models, integrating Anthropic models into Microsoft products (enterprise or consumer) requires CELA review, Zero Data Retention (ZDR), and must provide substantial additional functionality beyond what the Anthropic model provides alone.",
          considerations: [
            "CELA review is mandatory before any product integration",
            "ZDR (Zero Data Retention) must be enabled for all product scenarios",
            "The integration must add 'substantial additional functionality' — it cannot simply wrap the Anthropic model",
            "Trademark and branding rules apply (Section C.7) — Anthropic prior approval may be needed for logo use",
          ],
          nextSteps: [
            "Submit to the Third Party AI Models CELA Review site",
            "Contact Steven Stanford (sstanford@microsoft.com) for ZDR setup",
            "Document the 'substantial additional functionality' your product provides",
            "Review Section C requirements in full before beginning integration work",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section C: Product Integration Requirements", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Microsoft Responsible AI overview", url: MS_RESPONSIBLE_AI_URL },
            { text: "Govern AI — Cloud Adoption Framework", url: MS_AI_GOVERNANCE_URL },
          ],
        });
      }

      // ── Customer / consumer data (Section C.2) ──
      else if (
        q.includes("customer data") ||
        q.includes("consumer data") ||
        q.includes("user data") ||
        (q.includes("customer") && q.includes("data")) ||
        (q.includes("pii") && q.includes("anthropic"))
      ) {
        resolve({
          verdict: "Hold on — proceed with caution!",
          verdictType: "caution",
          explanation:
            "Per Section C.2 of the Engineering Guidance for Using Anthropic Models, using Anthropic models with customer or consumer data requires CELA review, privacy review, and compliance approvals before proceeding. This is a multi-step approval process.",
          considerations: [
            "CELA review is mandatory and must be completed before any customer data touches Anthropic models",
            "Privacy review is separately required for customer/consumer data scenarios",
            "Compliance team approval is needed in addition to CELA and privacy",
            "ZDR (Zero Data Retention) is mandatory for all customer data scenarios",
          ],
          nextSteps: [
            "Submit to the Third Party AI Models CELA Review site",
            "Initiate a privacy review through the standard privacy review process",
            "Contact your compliance team for scenario-specific approval",
            "Contact Steven Stanford (sstanford@microsoft.com) for ZDR setup",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section C.2: Customer & Consumer Data", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Microsoft Purview information protection", url: MS_PURVIEW_INFO_PROTECTION_URL },
            { text: "Microsoft Purview data security for generative AI apps", url: MS_PURVIEW_AI_URL },
          ],
        });
      }

      // ── EU / EFTA / UK enterprise (Section D.8) ──
      else if (
        q.includes("eu") ||
        q.includes("efta") ||
        q.includes("europe") ||
        q.includes("gdpr") ||
        q.includes("eudb") ||
        (q.includes("uk") && (q.includes("enterprise") || q.includes("data")))
      ) {
        resolve({
          verdict: "Hold on — proceed with caution!",
          verdictType: "caution",
          explanation:
            "Per Section D.8 of the Engineering Guidance for Using Anthropic Models, Anthropic model features must be OFF by default for E+D (Enterprise and Developer) offerings under Rajesh Jha's organization when serving EU/EFTA/UK customers. EUDB (EU Data Boundary) requirements cannot be met by Anthropic.",
          considerations: [
            "Anthropic models must be opt-in only (OFF by default) for EU/EFTA/UK enterprise customers",
            "EUDB compliance is not available with Anthropic — data will leave the EU boundary",
            "This restriction specifically applies to E+D offerings under Rajesh Jha's organization",
            "Consumer scenarios in these regions may have different requirements — check with CELA",
          ],
          nextSteps: [
            "Submit to the Third Party AI Models CELA Review site for region-specific guidance",
            "Ensure your feature design defaults Anthropic models to OFF for EU/EFTA/UK users",
            "Review Section D.8 for the complete list of regional requirements",
            "Consider Azure OpenAI as an alternative that meets EUDB requirements",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section D.8: EU/EFTA/UK Requirements", url: ANTHROPIC_GUIDANCE_URL },
            { text: "General Data Protection Regulation (GDPR)", url: MS_GDPR_URL },
            { text: "Govern AI apps and data for regulatory compliance", url: MS_AI_REGULATORY_URL },
          ],
        });
      }

      // ── Enterprise services — M365, Azure (Section D) ──
      else if (
        q.includes("enterprise") ||
        q.includes("m365") ||
        q.includes("microsoft 365") ||
        (q.includes("azure") && q.includes("anthropic")) ||
        q.includes("admin control")
      ) {
        resolve({
          verdict: "It depends on the details.",
          verdictType: "depends",
          explanation:
            "Per Section D of the Engineering Guidance for Using Anthropic Models, using Anthropic models in enterprise services (M365, Azure, etc.) has mandatory requirements: Zero Data Retention (ZDR), admin controls, API-only access, and no Government cloud or EUDB support.",
          considerations: [
            "ZDR (Zero Data Retention) is mandatory for all enterprise service scenarios",
            "IT admin controls must be provided for enterprise customers to manage Anthropic model access",
            "Only API-based integration is permitted — no direct UI embedding of Anthropic experiences",
            "Government cloud and EUDB environments are excluded (see Sections D.6 and D.8)",
          ],
          nextSteps: [
            "Submit to the Third Party AI Models CELA Review site",
            "Contact Steven Stanford (sstanford@microsoft.com) for ZDR setup",
            "Design admin controls for IT administrators to enable/disable the feature",
            "Review the full Section D requirements before architecture design",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section D: Enterprise Service Requirements", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Govern AI — Cloud Adoption Framework", url: MS_AI_GOVERNANCE_URL },
            { text: "Microsoft Purview data security for generative AI apps", url: MS_PURVIEW_AI_URL },
          ],
        });
      }

      // ── Synthetic data generation (Section E) ──
      else if (
        q.includes("synthetic data") ||
        q.includes("generate data") ||
        q.includes("data generation") ||
        (q.includes("synthetic") && q.includes("anthropic"))
      ) {
        resolve({
          verdict: "It depends on the details.",
          verdictType: "depends",
          explanation:
            "Per Section E of the Engineering Guidance for Using Anthropic Models, synthetic data generation using Anthropic models is permitted for benchmarking and evaluation purposes, but is prohibited if the synthetic data will be used for distillation or training competing models.",
          considerations: [
            "Synthetic data for benchmarking, testing, and evaluation is allowed",
            "Synthetic data for training or fine-tuning competing models is prohibited (Section A)",
            "The intended use of the generated data determines whether it's permitted",
            "Document the purpose of synthetic data generation for compliance records",
          ],
          nextSteps: [
            "Clearly document the intended use case for the synthetic data",
            "If for evaluation/benchmarking: proceed following standard procurement process",
            "If for model training: stop — this is prohibited under Section A",
            "Contact CELA if the intended use is ambiguous or novel",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section E: Synthetic Data Generation", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Anthropic Engineering Guidance — Section A: Prohibited Uses (distillation)", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Microsoft Responsible AI overview", url: MS_RESPONSIBLE_AI_URL },
          ],
        });
      }

      // ── Trademark / logo usage (Section C.7) ──
      else if (
        q.includes("trademark") ||
        q.includes("logo") ||
        q.includes("brand") ||
        (q.includes("anthropic") && q.includes("name"))
      ) {
        resolve({
          verdict: "Hold on — proceed with caution!",
          verdictType: "caution",
          explanation:
            "Per Section C.7 of the Engineering Guidance for Using Anthropic Models, use of Anthropic trademarks, logos, or branding in Microsoft products or materials requires prior written approval from Anthropic.",
          considerations: [
            "You cannot use the Anthropic or Claude logo without explicit approval",
            "Marketing materials referencing Anthropic must be reviewed",
            "Co-branding arrangements require formal agreement between Microsoft and Anthropic",
            "Internal documentation may have different requirements than external-facing materials",
          ],
          nextSteps: [
            "Submit a trademark usage request to Anthropic for approval",
            "Contact CELA for guidance on co-branding and trademark usage",
            "Review Section C.7 for the complete trademark usage requirements",
            "Use generic descriptions (e.g., 'third-party AI model') until approval is obtained",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Section C.7: Trademark & Branding", url: ANTHROPIC_GUIDANCE_URL },
          ],
        });
      }

      // ── Fallback / general (point to CELA review and procurement) ──
      else {
        resolve({
          verdict: "It depends on the details.",
          verdictType: "depends",
          explanation:
            "Your scenario doesn't match a specific category in the Engineering Guidance for Using Anthropic Models, but most Anthropic usage at Microsoft requires going through the CELA review process and approved procurement channels. The answer depends on your specific use case details.",
          considerations: [
            "Most Anthropic model usage requires CELA review — submit your scenario for assessment",
            "The procurement process must be followed for all Anthropic model access",
            "Zero Data Retention (ZDR) is required for any sensitive data scenarios",
            "Check which section (A–E) of the guidance most closely matches your use case",
          ],
          nextSteps: [
            "Submit to the Third Party AI Models CELA Review site with your specific scenario",
            "Ensure your team has completed the Anthropic procurement process",
            "Contact Steven Stanford (sstanford@microsoft.com) for ZDR and procurement questions",
            "Review the full Engineering Guidance for Using Anthropic Models document for details",
          ],
          sources: [
            { text: "Anthropic Engineering Guidance — Sections A–E (general guidance)", url: ANTHROPIC_GUIDANCE_URL },
            { text: "Govern AI — Cloud Adoption Framework", url: MS_AI_GOVERNANCE_URL },
            { text: "Microsoft Responsible AI overview", url: MS_RESPONSIBLE_AI_URL },
            { text: "Govern AI apps and data for regulatory compliance", url: MS_AI_REGULATORY_URL },
          ],
        });
      }
    }, delay);
  });
}

// ── Loading Messages ──

const LOADING_MESSAGES = [
  "Adjusting my glasses...",
  "Consulting the AI handbook...",
  "Running the scenario matrix...",
  "Cross-referencing policies...",
  "Hmm, let me think about this...",
  "Almost got it figured out...",
];

// ── App Component ──

function App() {
  const { instance, inProgress, accounts } = useMsal();
  const [scenario, setScenario] = useState("");
  const [response, setResponse] = useState<AIResponse | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [debugInfo, setDebugInfo] = useState("");

  const activeAccount = accounts[0] || null;

  useEffect(() => {
    const config = instance.getConfiguration();
    setDebugInfo(
      `Client ID: ${config.auth.clientId}\nRedirect URI: ${config.auth.redirectUri}\nAuthority: ${config.auth.authority}`
    );

    const params = new URLSearchParams(window.location.search);
    const hashParams = new URLSearchParams(
      window.location.hash.replace("#", "?")
    );
    const errorCode = params.get("error") || hashParams.get("error");
    const errorDesc =
      params.get("error_description") || hashParams.get("error_description");
    if (errorCode) {
      setError(`Azure AD Error: ${errorCode}\n${errorDesc || ""}`);
    }
  }, [instance]);

  const handleLogin = async () => {
    if (inProgress !== InteractionStatus.None) return;
    setError(null);
    try {
      await instance.loginRedirect(loginRequest);
    } catch (e: unknown) {
      const err = e as {
        errorCode?: string;
        errorMessage?: string;
        message?: string;
      };
      const errorMsg = err.errorCode
        ? `${err.errorCode}: ${err.errorMessage}`
        : err.message || "Sign-in failed";
      setError(errorMsg);
    }
  };

  const handleLogout = () => {
    instance.logoutRedirect();
  };

  const handleSubmit = useCallback(async () => {
    if (!scenario.trim() || isLoading) return;

    setIsLoading(true);
    setResponse(null);

    let msgIndex = 0;
    setLoadingMsg(LOADING_MESSAGES[0]);
    const interval = setInterval(() => {
      msgIndex = (msgIndex + 1) % LOADING_MESSAGES.length;
      setLoadingMsg(LOADING_MESSAGES[msgIndex]);
    }, 1100);

    try {
      const result = await evaluateScenario(scenario);
      setResponse(result);
    } finally {
      clearInterval(interval);
      setIsLoading(false);
    }
  }, [scenario, isLoading]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      handleSubmit();
    }
  };

  const handleNewQuestion = () => {
    setScenario("");
    setResponse(null);
  };

  const getInitials = (name?: string) => {
    if (!name) return "?";
    return name
      .split(" ")
      .map((n) => n[0])
      .join("")
      .toUpperCase()
      .slice(0, 2);
  };

  const verdictIcons: Record<VerdictType, string> = {
    yes: "\u2705",
    caution: "\u26A0\uFE0F",
    no: "\u274C",
    depends: "\uD83E\uDD14",
  };

  return (
    <>
      {/* ── Login Page ── */}
      <UnauthenticatedTemplate>
        <div className="login-page">
          <div className="login-card">
            <div className="login-glasses">
              <GlassesIcon size={100} />
            </div>

            <h1 className="login-title">Can I Do That?</h1>
            <p className="login-subtitle">Your AI Scenario Advisor</p>
            <p className="login-tagline">
              "Did I do that?" No — but <em>should</em> you? Let's find out.
            </p>

            <button
              className="ms-login-btn"
              onClick={handleLogin}
              disabled={inProgress !== InteractionStatus.None}
            >
              <MicrosoftLogo />
              {inProgress !== InteractionStatus.None
                ? "Signing in..."
                : "Sign in with Microsoft"}
            </button>

            {error && <div className="login-error">{error}</div>}

            {debugInfo && (
              <div className="login-debug">
                <strong>Azure AD Config:</strong>
                {debugInfo}
              </div>
            )}
          </div>
        </div>
      </UnauthenticatedTemplate>

      {/* ── Authenticated App ── */}
      <AuthenticatedTemplate>
        <div className="app-layout">
          {/* Header */}
          <header className="app-header">
            <div className="header-left">
              <div className="header-glasses">
                <GlassesIcon size={36} />
              </div>
              <span className="header-title">Can I Do That?</span>
            </div>
            <div className="header-right">
              <div className="user-info">
                <div className="user-avatar">
                  {getInitials(
                    activeAccount?.name || activeAccount?.username
                  )}
                </div>
                <span className="user-name">
                  {activeAccount?.name || activeAccount?.username}
                </span>
              </div>
              <button className="sign-out-btn" onClick={handleLogout}>
                Sign out
              </button>
            </div>
          </header>

          {/* Main */}
          <main className="app-main">
            <div className="welcome-section">
              <h2 className="welcome-title">Got an AI scenario?</h2>
              <p className="welcome-subtitle">
                Describe your Anthropic model usage scenario and I'll check it
                against the engineering guidance.
              </p>
            </div>

            {/* Question Input */}
            <div className="question-card">
              <label className="question-label">
                <span>💡</span> Describe your scenario
              </label>
              <textarea
                className="question-textarea"
                placeholder="e.g., I want to use Claude Code for internal development, or integrate Anthropic Claude into a Microsoft product..."
                value={scenario}
                onChange={(e) => setScenario(e.target.value)}
                onKeyDown={handleKeyDown}
                rows={4}
                maxLength={1000}
              />
              <div className="question-actions">
                <span className="char-count">
                  {scenario.length}/1000
                  {scenario.length > 0 && " · Ctrl+Enter to submit"}
                </span>
                <button
                  className="submit-btn"
                  onClick={handleSubmit}
                  disabled={!scenario.trim() || isLoading}
                >
                  Can I Do That? 🤓
                </button>
              </div>

              {/* Example Chips */}
              <div className="examples-section">
                <div className="examples-label">Try an example</div>
                <div className="example-chips">
                  {EXAMPLE_QUESTIONS.map((q, i) => (
                    <button
                      key={i}
                      className="example-chip"
                      onClick={() => setScenario(q)}
                    >
                      {q}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* Loading State */}
            {isLoading && (
              <div className="loading-card">
                <div className="loading-glasses">
                  <GlassesIcon size={48} />
                </div>
                <p className="loading-text">{loadingMsg}</p>
                <p className="loading-subtext">
                  Analyzing your AI scenario...
                </p>
                <div className="loading-bar">
                  <div className="loading-bar-fill" />
                </div>
              </div>
            )}

            {/* Response Card */}
            {response && !isLoading && (
              <div
                className={`response-card verdict-${response.verdictType}`}
              >
                <div className="verdict-banner">
                  <span className="verdict-icon">
                    {verdictIcons[response.verdictType]}
                  </span>
                  <span className="verdict-text">{response.verdict}</span>
                </div>
                <div className="response-body">
                  <p className="response-explanation">
                    {response.explanation}
                  </p>

                  {response.considerations.length > 0 && (
                    <div className="response-section">
                      <h3 className="response-section-title">
                        Things to Consider
                      </h3>
                      <ul className="response-list">
                        {response.considerations.map((c, i) => (
                          <li key={i}>{c}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  {response.nextSteps.length > 0 && (
                    <div className="response-section">
                      <h3 className="response-section-title">Next Steps</h3>
                      <ul className="response-list">
                        {response.nextSteps.map((s, i) => (
                          <li key={i}>{s}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  {response.sources.length > 0 && (
                    <div className="response-section">
                      <h3 className="response-section-title">Sources</h3>
                      <ul className="response-list">
                        {response.sources.map((s, i) => (
                          <li key={i}>
                            <a
                              href={s.url}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {s.text}
                            </a>
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}

                  <div className="response-actions">
                    <button
                      className="response-cta response-cta-primary"
                      onClick={handleNewQuestion}
                    >
                      🔍 Ask another question
                    </button>
                    <button
                      className="response-cta"
                      onClick={() => {
                        navigator.clipboard.writeText(
                          `Scenario: ${scenario}\n\nVerdict: ${response.verdict}\n\n${response.explanation}\n\nConsiderations:\n${response.considerations.map((c) => `• ${c}`).join("\n")}\n\nNext Steps:\n${response.nextSteps.map((s) => `• ${s}`).join("\n")}\n\nSources:\n${response.sources.map((s) => `• ${s.text}: ${s.url}`).join("\n")}`
                        );
                      }}
                    >
                      📋 Copy response
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Empty State */}
            {!response && !isLoading && !scenario && (
              <div className="empty-state">
                <div className="empty-glasses">
                  <GlassesIcon size={56} />
                </div>
                <p className="empty-text">No scenario yet? No problem!</p>
                <p className="empty-subtext">
                  Type your AI question above, or click an example to get
                  started.
                </p>
              </div>
            )}
          </main>

          {/* Footer */}
          <footer className="app-footer">
            <p className="footer-catchphrase">
              "Did I do that?" — Steve Urkel, and also you, probably
            </p>
            <p>Can I Do That? · AI Scenario Advisor · Microsoft Internal</p>
          </footer>
        </div>
      </AuthenticatedTemplate>
    </>
  );
}

export default App;
