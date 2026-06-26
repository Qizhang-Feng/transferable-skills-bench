"""
LLM client wrapper for the perturbation agent.
Uses ClaudeInference for cross-account Bedrock access.
"""

import json
import sys
import threading


_lock = threading.Lock()
_claude = None


def get_claude():
    global _claude
    if _claude is None:
        with _lock:
            if _claude is None:
                sys.path.insert(0, "vendor/PowerComputeCustomImage/configuration/rewards/Inference/APIInference")
                from api_invocation import ClaudeInference
                _claude = ClaudeInference()
    return _claude


def call_llm(prompt, system="You are a helpful assistant.", model="us.anthropic.claude-opus-4-6-v1",
             max_tokens=2000, temperature=0):
    """Call LLM and return raw text response."""
    claude = get_claude()
    prompt_dict = {"system": system, "user": prompt}
    return claude.inference(
        prompt_dict=prompt_dict,
        retry=3,
        guardrails=[],
        model_id=model,
        max_tokens=max_tokens,
        temperature=temperature,
    )


def call_llm_json(prompt, system="Respond only with valid JSON.", model="us.anthropic.claude-opus-4-6-v1",
                   max_tokens=2000, temperature=0):
    """Call LLM and parse JSON response."""
    text = call_llm(prompt, system=system, model=model,
                    max_tokens=max_tokens, temperature=temperature)

    # Extract JSON from response
    if "```json" in text:
        json_str = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        json_str = text.split("```")[1].split("```")[0].strip()
    else:
        json_str = text.strip()

    return json.loads(json_str)
