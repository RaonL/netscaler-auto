import os
import requests
import urllib3
import time

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

NETSCALERNODE1 = "x.x.x.211"
NEW_PASSWORD = os.getenv("NEW_PASSWORD")
if not NEW_PASSWORD:
    raise RuntimeError("NEW_PASSWORD environment variable not set")

CERTNAME = "star.namutech.co.kr_2024"

SF_VIPNAME = "SF_VIP"
SF_VIP = "x.x.x.219"
SF_SERVER1_NAME = "SF1_SRV"
SF_SERVER1_IP = "x.x.x.28"
SF_STORENAME = "ywlee"

LDAP_SERVER1_NAME = "AD1_SRV"
LDAP_SERVER1_IP = "x.x.x.20"
LDAP_SVC_ACCOUNT = "administrator@namulab.com"
LDAP_SVC_PASSWORD = os.getenv("LDAP_SVC_PASSWORD")
if not LDAP_SVC_PASSWORD:
    raise RuntimeError("LDAP_SVC_PASSWORD environment variable not set")
LDAP_BASE_DN = "DC=namulab,DC=com"

CTXGW_VIPNAME = "CTXGW_VIP"
CTXGW_VIP = os.getenv("CTXGW_VIP")
if not CTXGW_VIP:
    raise RuntimeError("CTXGW_VIP environment variable not set")
STA1 = "https://x.x.x.23"

SSO_DOMAIN = "namulab"

class NetScalerConfigurator:
    def __init__(self, node, password):
        self.base = f"https://{node}"
        self.session = requests.Session()
        self.session.verify = False
        self.session.headers.update({
            "X-NITRO-USER": "nsroot",
            "X-NITRO-PASS": password
        })

    def request(self, method, endpoint, data=None, desc="", ignore_error=False):
        print(f"  -> {desc}")
        try:
            resp = self.session.request(method, self.base + endpoint, json=data)
            resp.raise_for_status()
            print("     ✓ 성공")
            if resp.text:
                try:
                    return resp.json()
                except ValueError:
                    return resp.text
            return None
        except requests.RequestException as e:
            if ignore_error:
                print(f"     ⚠ 무시됨: {e}")
            else:
                print(f"     ✗ 실패: {e}")
                if e.response is not None:
                    print(f"     응답: {e.response.text}")
            return None

def main():
    ns = NetScalerConfigurator(NETSCALERNODE1, NEW_PASSWORD)

    print("\n=== 기존 구성 정리 (철저한 삭제) ===")

    possible_ldap_policies = ["LDAP_POL", "x.x.x.20_LDAP_pol", "LDAP_Policy", "ldap_pol", "BASIC_LDAP_POL", "BASIC_LDAP"]
    for policy in possible_ldap_policies:
        ns.request("DELETE", f"/nitro/v1/config/authenticationpolicy/{policy}", desc=f"{policy} Advanced 정책 삭제", ignore_error=True)
        ns.request("DELETE", f"/nitro/v1/config/authenticationldappolicy/{policy}", desc=f"{policy} Basic 정책 삭제", ignore_error=True)

    possible_ldap_actions = ["LDAP_ACT", "LDAP_Action", "ldap_act"]
    for action in possible_ldap_actions:
        ns.request("DELETE", f"/nitro/v1/config/authenticationldapaction/{action}", desc=f"{action} 액션 삭제", ignore_error=True)

    bindings = [
        "vpnvserver_authenticationpolicy_binding/CTXGW_VIP?policy=LDAP_POL",
        "vpnvserver_authenticationldappolicy_binding/CTXGW_VIP?policy=LDAP_POL",
        "vpnvserver_authenticationldappolicy_binding/CTXGW_VIP?policy=BASIC_LDAP_POL",
        "vpnvserver_authenticationldappolicy_binding/CTXGW_VIP?policy=BASIC_LDAP",
        "vpnvserver_vpnsessionpolicy_binding/CTXGW_VIP?policy=Web_POL",
        "vpnvserver_vpnsessionpolicy_binding/CTXGW_VIP?policy=Workspace_App_POL",
        "vpnvserver_staserver_binding/CTXGW_VIP?staserver=https://10.10.11.23",
        f"sslvserver_sslcertkey_binding/CTXGW_VIP?certkeyname={CERTNAME}",
        "lbvserver_servicegroup_binding/SF_VIP?servicegroupname=SF_SVG",
        "servicegroup_servicegroupmember_binding/SF_SVG?servername=SF1_SRV",
        "servicegroup_lbmonitor_binding/SF_SVG?monitor_name=SF_MON",
        f"sslvserver_sslcertkey_binding/SF_VIP?certkeyname={CERTNAME}"
    ]

    for binding in bindings:
        ns.request("DELETE", f"/nitro/v1/config/{binding}", desc=f"바인딩 해제: {binding}", ignore_error=True)

    items_to_delete = [
        {"type": "vpnvserver", "name": "CTXGW_VIP"},
        {"type": "vpnsessionpolicy", "name": "Web_POL"},
        {"type": "vpnsessionpolicy", "name": "Workspace_App_POL"},
        {"type": "vpnsessionaction", "name": "Web_ACT"},
        {"type": "vpnsessionaction", "name": "Workspace_App_ACT"},
        {"type": "authenticationpolicy", "name": "LDAP_POL"},
        {"type": "authenticationldapaction", "name": "LDAP_ACT"},
        {"type": "lbvserver", "name": "SF_VIP"},
        {"type": "lbmonitor", "name": "SF_MON"},
        {"type": "servicegroup", "name": "SF_SVG"},
        {"type": "server", "name": "SF1_SRV"},
        {"type": "server", "name": "AD1_SRV"}
    ]

    for item in items_to_delete:
        ns.request("DELETE", f"/nitro/v1/config/{item['type']}/{item['name']}", desc=f"{item['name']} 삭제", ignore_error=True)

    time.sleep(3)

    print("\n=== StoreFront 로드밸런서 구성 ===")

    ns.request("POST", "/nitro/v1/config/server", data={"server": {"name": SF_SERVER1_NAME, "ipaddress": SF_SERVER1_IP}}, desc="StoreFront 서버 생성")
    ns.request("POST", "/nitro/v1/config/servicegroup", data={"servicegroup": {"servicegroupname": "SF_SVG", "servicetype": "SSL"}}, desc="StoreFront ServiceGroup 생성")
    ns.request("POST", "/nitro/v1/config/servicegroup_servicegroupmember_binding", data={"servicegroup_servicegroupmember_binding": {"servername": SF_SERVER1_NAME, "servicegroupname": "SF_SVG", "port": 443}}, desc="StoreFront 서버를 ServiceGroup에 바인딩")
    ns.request("POST", "/nitro/v1/config/lbmonitor", data={"lbmonitor": {"monitorname": "SF_MON", "type": "STOREFRONT", "secure": "YES", "storename": SF_STORENAME}}, desc="StoreFront 모니터 생성")
    ns.request("POST", "/nitro/v1/config/servicegroup_lbmonitor_binding", data={"servicegroup_lbmonitor_binding": {"servicegroupname": "SF_SVG", "monitor_name": "SF_MON"}}, desc="StoreFront 모니터를 ServiceGroup에 바인딩")
    ns.request("POST", "/nitro/v1/config/lbvserver", data={"lbvserver": {"name": SF_VIPNAME, "servicetype": "SSL", "ipv46": SF_VIP, "port": 443, "persistencetype": "COOKIEINSERT", "timeout": 60}}, desc="StoreFront LB 가상서버 생성")
    ns.request("POST", "/nitro/v1/config/lbvserver_servicegroup_binding", data={"lbvserver_servicegroup_binding": {"servicegroupname": "SF_SVG", "name": SF_VIPNAME}}, desc="ServiceGroup을 StoreFront LB에 바인딩")
    ns.request("POST", "/nitro/v1/config/sslvserver_sslcertkey_binding", data={"sslvserver_sslcertkey_binding": {"vservername": SF_VIPNAME, "certkeyname": CERTNAME}}, desc="StoreFront LB에 SSL 인증서 바인딩")

    print("\n=== LDAP 서버 구성 (직접 연결) ===")
    ns.request("POST", "/nitro/v1/config/server", data={"server": {"name": LDAP_SERVER1_NAME, "ipaddress": LDAP_SERVER1_IP}}, desc="AD 서버 생성 (직접 연결용)")

    print("\n=== LDAP 인증 구성 (공식 문서 기준) ===")
    action_body = {
        "authenticationldapaction": {
            "name": "LDAP_ACT",
            "serverip": LDAP_SERVER1_IP,
            "sectype": "SSL",
            "serverport": "636",
            "authtimeout": "3",
            "authentication": "ENABLED",
            "ldapbase": LDAP_BASE_DN,
            "ldapbinddn": LDAP_SVC_ACCOUNT,
            "ldapbinddnpassword": LDAP_SVC_PASSWORD,
            "ldaploginname": "sAMAccountName",
            "email": "mail",
            "requireuser": "YES",
            "passwdchange": "ENABLED",
            "ssonameattribute": "UserPrincipalName",
            "groupattrname": "memberOf",
            "subattributename": "cn"
        }
    }

    resp = ns.request("POST", "/nitro/v1/config/authenticationldapaction", data=action_body, desc="LDAP 인증 액션 생성 (LDAP_ACT)")
    action_success = resp is not None
    time.sleep(3)

    global_policy = None
    if action_success:
        policy_body = {
            "authenticationldappolicy": {
                "name": "LDAP_POL",
                "rule": "NS_TRUE",
                "reqaction": "LDAP_ACT"
            }
        }
        resp = ns.request("POST", "/nitro/v1/config/authenticationldappolicy", data=policy_body, desc="LDAP Policy 생성 (Basic Authentication - NS_TRUE 규칙)")
        if resp is not None:
            global_policy = "LDAP_POL"
            print("  ✓ LDAP 정책 생성 성공: LDAP_POL (Basic Authentication)")
        else:
            print("  ✗ LDAP 정책 생성 실패")
    else:
        print("  ✗ LDAP 액션 생성 실패로 정책 생성 건너뜀")

    print("\n=== VPN 세션 구성 ===")
    ns.request("POST", "/nitro/v1/config/vpnsessionaction", data={"vpnsessionaction": {"name": "Web_ACT", "sesstimeout": 60, "transparentinterception": "OFF", "defaultauthorizationaction": "ALLOW", "icaproxy": "ON", "wihome": f"https://{SF_VIP}/Citrix/{SF_STORENAME}Web"}}, desc="웹 세션 액션 생성 (기본 설정)")
    ns.request("POST", "/nitro/v1/config/vpnsessionaction", data={"vpnsessionaction": {"name": "Workspace_App_ACT", "sesstimeout": 60, "transparentinterception": "OFF", "defaultauthorizationaction": "ALLOW", "icaproxy": "ON", "storefronturl": f"https://{SF_VIP}"}}, desc="Workspace App 세션 액션 생성 (기본 설정)")
    ns.request("POST", "/nitro/v1/config/vpnsessionpolicy", data={"vpnsessionpolicy": {"name": "Web_POL", "action": "Web_ACT", "rule": 'HTTP.REQ.HEADER("User-Agent").CONTAINS("CitrixReceiver").NOT'}}, desc="웹 세션 정책 생성")
    ns.request("POST", "/nitro/v1/config/vpnsessionpolicy", data={"vpnsessionpolicy": {"name": "Workspace_App_POL", "action": "Workspace_App_ACT", "rule": 'HTTP.REQ.HEADER("User-Agent").CONTAINS("CitrixReceiver")'}}, desc="Workspace App 세션 정책 생성")

    print("\n=== Citrix Gateway 구성 ===")
    ns.request("POST", "/nitro/v1/config/vpnvserver", data={"vpnvserver": {"name": CTXGW_VIPNAME, "servicetype": "SSL", "ipv46": CTXGW_VIP, "port": 443, "maxloginattempts": 5, "failedlogintimeout": 15, "icaonly": "ON", "dtls": "OFF"}}, desc="Citrix Gateway 가상서버 생성 (ICA Only ON, DTLS OFF)")
    ns.request("POST", "/nitro/v1/config/sslvserver_sslcertkey_binding", data={"sslvserver_sslcertkey_binding": {"vservername": CTXGW_VIPNAME, "certkeyname": CERTNAME}}, desc="Gateway에 SSL 인증서 바인딩")
    time.sleep(3)

    print("\n=== LDAP 정책 바인딩 (Basic Authentication) ===")
    if global_policy:
        ns.request("POST", "/nitro/v1/config/vpnvserver_authenticationldappolicy_binding", data={"vpnvserver_authenticationldappolicy_binding": {"name": CTXGW_VIPNAME, "policy": global_policy, "priority": 100}}, desc=f"Basic LDAP 인증 정책 ({global_policy})을 Gateway에 바인딩")
    else:
        print("  ✗ LDAP 정책이 생성되지 않음 - 수동으로 생성 필요")

    ns.request("POST", "/nitro/v1/config/vpnvserver_vpnsessionpolicy_binding", data={"vpnvserver_vpnsessionpolicy_binding": {"name": CTXGW_VIPNAME, "policy": "Web_POL", "priority": 100}}, desc="웹 세션 정책을 Gateway에 바인딩")
    ns.request("POST", "/nitro/v1/config/vpnvserver_vpnsessionpolicy_binding", data={"vpnvserver_vpnsessionpolicy_binding": {"name": CTXGW_VIPNAME, "policy": "Workspace_App_POL", "priority": 110}}, desc="Workspace App 세션 정책을 Gateway에 바인딩")
    ns.request("POST", "/nitro/v1/config/vpnvserver_staserver_binding", data={"vpnvserver_staserver_binding": {"name": CTXGW_VIPNAME, "staserver": STA1}}, desc="STA 서버를 Gateway에 바인딩")

    print("\n=== 세션 프로파일 업데이트 (공식 문서 기준) ===")
    ns.request("PUT", "/nitro/v1/config/vpnsessionaction", data={"vpnsessionaction": {"name": "Web_ACT", "sso": "ON", "ntdomain": SSO_DOMAIN, "httpport": [80]}}, desc="Web_ACT에 SSO 및 NT Domain 설정")
    ns.request("PUT", "/nitro/v1/config/vpnsessionaction", data={"vpnsessionaction": {"name": "Workspace_App_ACT", "sso": "ON", "ntdomain": SSO_DOMAIN}}, desc="Workspace_App_ACT에 SSO 및 NT Domain 설정")

    print("  단계 3: 설정 확인")
    for profile in ["Web_ACT", "Workspace_App_ACT"]:
        resp = ns.request("GET", f"/nitro/v1/config/vpnsessionaction/{profile}", desc=f"{profile} 확인", ignore_error=True)
        if resp and 'vpnsessionaction' in resp:
            action = resp['vpnsessionaction']
            print(f"    {profile} 설정:")
            print(f"      SSO: {action.get('sso')}")
            print(f"      NT Domain: {action.get('ntdomain')}")
            if 'httpport' in action:
                print(f"      HTTP Port: {','.join(map(str, action['httpport']))}")
            if 'wihome' in action:
                print(f"      Web Interface: {action['wihome']}")
            if 'storefronturl' in action:
                print(f"      StoreFront URL: {action['storefronturl']}")

    print("\n=== 구성 완료 ===")
    print(f"StoreFront URL: https://{SF_VIP}")
    print(f"LDAP 서버: {LDAP_SERVER1_IP}:636 (직접 연결)")
    print(f"Citrix Gateway URL: https://{CTXGW_VIP}")
    print(f"STA Server: {STA1}")
    print(f"SSO Domain: {SSO_DOMAIN}")

if __name__ == "__main__":
    main()
