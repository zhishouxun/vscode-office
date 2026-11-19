"""Version-bound changes to the 4.2.3 host's local feature policy.

No license is generated or written. The bundled controls use isPro as a
capability flag; local builds always expose their already bundled features.
Exact match counts deliberately reject incompatible upstream bundles.
"""
import re


def replace_checked(source, pattern, replacement, count):
    result, actual = re.subn(pattern, lambda _: replacement, source)
    if actual != count:
        raise ValueError(f'Unsupported host bundle: {pattern!r}: expected {count}, got {actual}')
    return result


def patch_host(source, web=False):
    service, predicate, key_reader, cancelled, status, pdf_gate, broadcast, vscode = (
        ('hr', 'Dw', 'Ds', '_f', 'mE', 'BN', 'Cw', 'No') if web else
        ('Tc', 'Kte', 'zA', 'U_', 'mU', 'iG', 'Xte', 'Y1')
    )
    source = replace_checked(source,
        re.escape(f'function {predicate}(t){{return!!{key_reader}(t)}}'),
        f'function {predicate}(t){{return true}}', 1)
    source = replace_checked(source,
        re.escape(f'function {cancelled}(t){{return {status}(t)==="cancel"}}'),
        f'function {cancelled}(t){{return false}}', 1)
    # Markdown and PDF status payloads. Leave the actual license reader intact.
    source = replace_checked(source,
        rf'!!{service}\.getLicenseKey\((?:this\.context|r)\)', 'true', 4)
    gate = (rf'function {pdf_gate}\(t,e\)\{{return {service}\.getLicenseKey\(t\)'
            rf'\?!0:\(e\.emit\("pdfTool\.result",\{{ok:!1,error:"Pro license required",needPro:!0\}}\),'
            rf'\w+\.commands\.executeCommand\("office\.pro\.activate"\),!1\)\}}')
    source = replace_checked(source, gate,
        f'function {pdf_gate}(t,e){{return true}}', 1)
    # Ignore account-state broadcasts in the local distribution.
    source = replace_checked(source, re.escape(f'function {broadcast}(t,e){{'),
        f'function {broadcast}(t,e){{e=true;', 1)
    # No license network refresh is needed by a local feature policy.
    source = replace_checked(source,
        re.escape(f'{service}.refreshLatestLicense(t)'), 'void 0', 1)
    # Retain the command ID for compatibility, with an honest local message.
    source = replace_checked(source,
        'static async createOrShow\\(e\\)\\{if\\(this.panel\\)',
        f'static async createOrShow(e){{return {vscode}.window.showInformationMessage('
        '"本地版已免费开放全部内置功能，无需激活。");if(this.panel)', 1)
    return source
