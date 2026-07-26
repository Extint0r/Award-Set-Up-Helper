from datetime import datetime
from typing import List, Dict, Any

def synthesize_portfolio_pass_2(cluster_docs: List[Dict[str, Any]], system_baselines: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    """Aggregates portfolio data and calculates financial audit verdicts per award cluster."""
    if not cluster_docs:
        return {}

    cluster_key = cluster_docs[0]["AWARD_CLUSTER_KEY"]
    
    oracle_num  = next((d["ORACLE_AWARD_NUMBER"] for d in cluster_docs if d["ORACLE_AWARD_NUMBER"]), "")
    cayuse_proj = next((d["CAYUSE_PROJECT_NUMBER"] for d in cluster_docs if d["CAYUSE_PROJECT_NUMBER"]), "")

    baseline = (
        system_baselines.get(oracle_num) or 
        system_baselines.get(cayuse_proj) or 
        system_baselines.get(cluster_key, {})
    )

    if not oracle_num:
        oracle_num = baseline.get("ORACLE_AWARD_NUMBER", "")
    if not cayuse_proj:
        cayuse_proj = baseline.get("CAYUSE_PROJECT_NUMBER", "")

    cayuse_ob      = baseline.get("CAYUSE_OBLIGATED", 0.0)
    oracle_ob      = baseline.get("ORACLE_OBLIGATED", 0.0)
    cayuse_ceiling = baseline.get("CAYUSE_TOTAL_AMOUNT", 0.0)
    oracle_ceiling = baseline.get("ORACLE_HEADER_HARD_LIMIT", 0.0)

    c_0 = max(cayuse_ceiling, oracle_ceiling)
    active_ceiling = c_0

    sorted_docs = sorted(
        cluster_docs, 
        key=lambda x: (x["EXECUTION_DATE"] or x["DOC_START_DATE"] or datetime.min)
    )

    start_dates = [d["DOC_START_DATE"] for d in sorted_docs if d["DOC_START_DATE"]]
    end_dates   = [d["DOC_END_DATE"] for d in sorted_docs if d["DOC_END_DATE"]]
    
    award_start_date_truth = min(start_dates) if start_dates else None
    award_end_date_truth   = max(end_dates) if end_dates else None
    
    cum_obligated = 0.0
    highest_doc_ceiling = 0.0
    
    for doc in sorted_docs:
        cum_obligated += doc["DELTA_OBLIGATED"]
        if doc["DOC_CEILING"] > 0:
            active_ceiling = doc["DOC_CEILING"]
            highest_doc_ceiling = max(highest_doc_ceiling, doc["DOC_CEILING"])

    if active_ceiling == 0.0:
        active_ceiling = c_0

    ceiling_breach = False
    audit_status = "IN_SYNC"
    
    if cum_obligated > active_ceiling and active_ceiling > 0:
        if highest_doc_ceiling >= cum_obligated:
            active_ceiling = highest_doc_ceiling
            audit_status = "CEILING_RESOLVED_BY_CLUSTER_DOC"
        else:
            ceiling_breach = True
            audit_status = "CEILING_BREACH_UNAPPROVED"

    cay_sync = (abs(cum_obligated - cayuse_ob) < 1.0)
    orc_sync = (abs(cum_obligated - oracle_ob) < 1.0)
    
    if ceiling_breach:
        reconciliation_verdict = "FLAG_CEILING_BREACH"
        discrepancy_reason = f"Cumulative PDF obligations (${cum_obligated:,.2f}) exceed active ceiling (${active_ceiling:,.2f})."
    elif cay_sync and orc_sync:
        reconciliation_verdict = "IN_SYNC"
        discrepancy_reason = "PDF Truth, Cayuse, and Oracle amounts are fully aligned."
    elif not cay_sync and orc_sync:
        reconciliation_verdict = "CAYUSE_UPDATE_REQ"
        discrepancy_reason = f"PDF Truth (${cum_obligated:,.2f}) matches Oracle, but Cayuse (${cayuse_ob:,.2f}) requires update."
    elif cay_sync and not orc_sync:
        reconciliation_verdict = "ORACLE_UPDATE_REQ"
        discrepancy_reason = f"PDF Truth (${cum_obligated:,.2f}) matches Cayuse, but Oracle (${oracle_ob:,.2f}) requires update."
    else:
        reconciliation_verdict = "BOTH_OUT_OF_SYNC"
        if cayuse_ob == 0.0 and oracle_ob == 0.0 and cum_obligated > 0:
            discrepancy_reason = f"UNMATCHED_BASELINE: Extracted PDF funds (${cum_obligated:,.2f}), but no Cayuse/Oracle baseline record was matched."
        elif cum_obligated == 0.0 and (cayuse_ob > 0 or oracle_ob > 0):
            discrepancy_reason = f"ZERO_PDF_EXTRACTION: Baseline has obligations (Cayuse: ${cayuse_ob:,.2f}, Oracle: ${oracle_ob:,.2f}), but $0 action delta extracted from PDFs."
        else:
            discrepancy_reason = f"FINANCIAL_DISCREPANCY: PDF Truth (${cum_obligated:,.2f}) differs from Cayuse (${cayuse_ob:,.2f}) and Oracle (${oracle_ob:,.2f})."

    return {
        "AWARD_CLUSTER_KEY": cluster_key,
        "ORACLE_AWARD_NUMBER": oracle_num,
        "CAYUSE_PROJECT_NUMBER": cayuse_proj,
        "DOCUMENT_COUNT": len(sorted_docs),
        "PDF_START_DATE_TRUTH": award_start_date_truth,
        "PDF_END_DATE_TRUTH": award_end_date_truth,
        "PDF_CUMULATIVE_OBLIGATED": cum_obligated,
        "PDF_ACTIVE_CEILING": active_ceiling,
        "CAYUSE_OBLIGATED": cayuse_ob,
        "ORACLE_OBLIGATED": oracle_ob,
        "CAYUSE_CEILING": cayuse_ceiling,
        "ORACLE_CEILING": oracle_ceiling,
        "CEILING_BREACH": ceiling_breach,
        "AUDIT_STATUS": audit_status,
        "RECONCILIATION_VERDICT": reconciliation_verdict,
        "DISCREPANCY_REASON": discrepancy_reason
    }