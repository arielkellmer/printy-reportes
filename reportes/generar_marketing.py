"""Printy - dashboard de marketing: visitas (GA4), origen de cada venta (WooCommerce
order attribution) y rendimiento por anuncio.

Uso:
    python generar_marketing.py [--dias 90]

Credenciales (env primero, archivo local como fallback):
    WC_CONSUMER_KEY / WC_CONSUMER_SECRET  -> WooCommerce REST (o data/datos.txt)
    GA_SA_JSON                            -> JSON de la cuenta de servicio GA4, texto o base64
                                             (o data/printy-seo-*.json)
    META_ACCESS_TOKEN / META_AD_ACCOUNT   -> opcional: nombres + inversión por anuncio de Meta
Salida: reportes/Marketing_Printy.html
"""
import argparse
import base64
import glob
import json
import os
import time
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from generar_reporte import ROOT, VALID_STATUSES, fetch_all, load_credentials

HERE = Path(__file__).resolve().parent
GA_PROPERTY = '378383956'
FIX_DATE = '2026-09-23'  # arreglo del tracking de Imaxel (GA4 + WooCommerce)
META_API = 'https://graph.facebook.com/v23.0'
AR_TZ = timezone(timedelta(hours=-3))

CHANNELS = [  # orden fijo = orden de color
    ('meta', 'Meta Ads', True),
    ('buscadores', 'Google y buscadores', False),
    ('directo', 'Directo', False),
    ('redes', 'Instagram/Facebook orgánico', False),
    ('tiktok', 'TikTok', False),
    ('whatsapp', 'WhatsApp / Chatbot', False),
    ('email', 'Email (Perfit)', False),
    ('otros', 'Otros', False),
    ('imaxel', 'Imaxel (origen perdido)', False),
]

META_SRC = {'ig', 'fb', 'an', 'th', 'msg', 'facebook', 'instagram', 'meta'}
PAID_MED = {'paid', 'paid_social', 'paidsocial', 'cpc', 'ppc', 'ads', 'paid-social'}
SEARCH = ('google', 'bing', 'yahoo', 'duckduckgo', 'ecosia', 'yandex', 'baidu', 'search')


def classify(src, med):
    """(canal, plataforma de publicidad o None) a partir de source/medium."""
    s = (src or '').lower().strip()
    m = (med or '').lower().strip()
    if 'imaxel' in s:
        return 'imaxel', None
    if m in PAID_MED:
        if s in META_SRC or 'facebook' in s or 'instagram' in s:
            return 'meta', 'Meta Ads'
        if 'tiktok' in s:
            return 'tiktok', 'TikTok Ads'
        if 'google' in s:
            return 'otros', 'Google Ads'
        return 'otros', None
    if s == 'ads.tiktok.com':
        return 'tiktok', 'TikTok Ads'
    if 'tiktok' in s:
        return 'tiktok', None
    if 'notchatbot' in s or 'whatsapp' in m or s in ('l.wl.co', 'wa.me', 'api.whatsapp.com'):
        return 'whatsapp', None
    if m == 'email' or 'perfit' in s:
        return 'email', 'Email'
    if m == 'organic' or any(k in s for k in SEARCH):
        return 'buscadores', None
    if s in META_SRC or 'instagram' in s or 'facebook' in s or m == 'social':
        return 'redes', None
    if s in ('(direct)', 'direct', '') and m in ('(none)', 'none', 'typein', ''):
        return 'directo', None
    return 'otros', None


def clean(v):
    v = (v or '').strip()
    return '' if v in ('(not set)', '(none)', '(direct)', '(organic)', '(referral)',
                       '(ai-assistant)', '(data not available)') else v


# ---------------------------------------------------------------- GA4
def ga_session():
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2 import service_account
    scopes = ['https://www.googleapis.com/auth/analytics.readonly',
              'https://www.googleapis.com/auth/webmasters.readonly']
    raw = os.environ.get('GA_SA_JSON')
    if raw:
        raw = raw.strip()
        if not raw.startswith('{'):
            raw = base64.b64decode(raw).decode('utf-8')
        cred = service_account.Credentials.from_service_account_info(json.loads(raw), scopes=scopes)
    else:
        path = sorted(glob.glob(str(ROOT / 'data' / 'printy-seo-*.json')))[0]
        cred = service_account.Credentials.from_service_account_file(path, scopes=scopes)
    return AuthorizedSession(cred)


def ga_report(sess, desde, hasta, dims, mets):
    rows, offset = [], 0
    while True:
        body = {
            'dateRanges': [{'startDate': desde, 'endDate': hasta}],
            'dimensions': [{'name': d} for d in dims],
            'metrics': [{'name': m} for m in mets],
            'limit': 100000, 'offset': offset,
        }
        r = sess.post(f'https://analyticsdata.googleapis.com/v1beta/properties/{GA_PROPERTY}:runReport', json=body)
        j = r.json()
        if 'error' in j:
            raise RuntimeError(j['error'].get('message'))
        batch = j.get('rows', [])
        for row in batch:
            rows.append([v['value'] for v in row['dimensionValues']] + [float(v['value']) for v in row['metricValues']])
        offset += len(batch)
        if not batch or offset >= j.get('rowCount', 0):
            return rows


def fetch_ga(desde, hasta):
    sess = ga_session()
    raw = ga_report(sess, desde, hasta,
                    ['date', 'sessionSource', 'sessionMedium', 'sessionCampaignName',
                     'sessionManualTerm', 'sessionManualAdContent'],
                    ['sessions', 'engagedSessions', 'addToCarts', 'ecommercePurchases', 'purchaseRevenue'])
    users = ga_report(sess, desde, hasta, ['date'], ['totalUsers', 'newUsers', 'sessions'])
    daily = defaultdict(lambda: [0, 0, 0, 0, 0.0])   # (date, canal) -> ses, eng, carts, compras, ingresos
    ads = defaultdict(lambda: [0, 0, 0, 0, 0.0])     # (date, plataforma, campaña, conjunto, anuncio)
    for d, src, med, camp, term, content, ses, eng, carts, purch, rev in raw:
        day = f'{d[:4]}-{d[4:6]}-{d[6:]}'
        ch, plat = classify(src, med)
        acc = daily[(day, ch)]
        for i, v in enumerate((ses, eng, carts, purch, rev)):
            acc[i] += v
        if plat:
            a = ads[(day, plat, clean(camp), clean(term), clean(content))]
            for i, v in enumerate((ses, eng, carts, purch, rev)):
                a[i] += v
    ga_daily = [[k[0], k[1], int(v[0]), int(v[1]), int(v[2]), int(v[3]), round(v[4])] for k, v in sorted(daily.items())]
    ga_ads = [[*k, int(v[0]), int(v[1]), int(v[2]), int(v[3]), round(v[4])] for k, v in sorted(ads.items())]
    ga_users = sorted([[f'{d[:4]}-{d[4:6]}-{d[6:]}', int(u), int(n), int(s)] for d, u, n, s in users])
    return ga_daily, ga_ads, ga_users


# ---------------------------------------------------------------- Search Console
GSC_SITE = 'https://printy.photos/'
BRAND = ('printy', 'megaphoto', 'mega photo', 'printi')


def url_path(u):
    p = urllib.parse.urlparse(u).path or '/'
    return p if p.endswith('/') else p + '/'


def fetch_gsc(desde, hasta):
    """Google no pasa la búsqueda de cada visita; Search Console da los totales por
    búsqueda y página. Devuelve (búsquedas diarias con clics, top búsquedas por página)."""
    sess = ga_session()
    url = ('https://searchconsole.googleapis.com/webmasters/v3/sites/'
           + urllib.parse.quote(GSC_SITE, safe='') + '/searchAnalytics/query')

    def query(dims):
        rows, start = [], 0
        while True:
            r = sess.post(url, json={'startDate': desde, 'endDate': hasta, 'dimensions': dims,
                                     'rowLimit': 25000, 'startRow': start, 'dataState': 'all'}).json()
            if 'error' in r:
                raise RuntimeError(r['error'].get('message'))
            batch = r.get('rows', [])
            rows += batch
            if len(batch) < 25000:
                return rows
            start += 25000

    daily = [[r['keys'][0], r['keys'][1], url_path(r['keys'][2]), int(r['clicks']), int(r['impressions'])]
             for r in query(['date', 'query', 'page']) if r['clicks'] > 0]
    by_page = defaultdict(list)
    for r in query(['page', 'query']):
        by_page[url_path(r['keys'][0])].append([r['keys'][1], int(r['clicks']), int(r['impressions']),
                                                round(r['position'], 1)])
    pages = {p: sorted(q, key=lambda x: (-x[1], -x[2]))[:5] for p, q in by_page.items()}
    last = max((d[0] for d in daily), default=hasta)
    return daily, pages, last


# ---------------------------------------------------------------- WooCommerce
def fetch_orders(desde, hasta):
    ck, cs = load_credentials()
    orders = fetch_all(ck, cs, 'orders', {'after': f'{desde}T00:00:00', 'before': f'{hasta}T23:59:59', 'status': 'any'})
    out = []
    for o in orders:
        if o.get('status') not in VALID_STATUSES:
            continue
        m = {x['key'][len('_wc_order_attribution_'):]: x['value'] for x in o.get('meta_data', [])
             if str(x.get('key', '')).startswith('_wc_order_attribution_')}
        st = m.get('source_type', '')
        src, med = m.get('utm_source', ''), m.get('utm_medium', '')
        if st == 'typein':
            src, med = '(direct)', '(none)'
        elif st == 'admin':
            src, med = 'pedido manual', 'admin'
        elif not st:
            src, med = 'sin dato', ''
        ch, plat = classify(src, med)
        entry = m.get('session_entry', '')
        try:
            entry = url_path(entry) if entry else ''
        except ValueError:
            pass
        items = '; '.join(f"{li.get('name', '')}{' x' + str(li['quantity']) if li.get('quantity', 1) > 1 else ''}"
                          for li in o.get('line_items', []))
        out.append({
            'id': o['id'], 'n': o.get('number', o['id']),
            'dt': o['date_created'][:16],
            'total': round(float(o.get('total') or 0)),
            'ch': ch, 'plat': plat or '',
            'src': src, 'med': med, 'type': st,
            'camp': m.get('utm_campaign', ''), 'adset': m.get('utm_term', ''), 'ad': m.get('utm_content', ''),
            'entry': entry, 'dev': m.get('device_type', ''),
            'pages': m.get('session_pages', ''), 'visits': m.get('session_count', ''),
            'items': items[:240],
        })
    out.sort(key=lambda r: r['dt'], reverse=True)
    return out


# ---------------------------------------------------------------- Meta (opcional)
def fetch_meta(desde, hasta):
    token, accounts = os.environ.get('META_ACCESS_TOKEN'), os.environ.get('META_AD_ACCOUNT')
    if not token or not accounts:
        return None
    names, spend = {}, []
    for account in (a.strip() for a in accounts.split(',') if a.strip()):  # admite varias cuentas separadas por coma
        account = account if account.startswith('act_') else f'act_{account}'
        params = {
            'level': 'ad', 'time_increment': 1, 'limit': 500, 'access_token': token,
            'fields': 'ad_id,ad_name,adset_id,adset_name,campaign_id,campaign_name,spend,impressions,clicks',
            'time_range': json.dumps({'since': desde, 'until': hasta}),
        }
        url = f'{META_API}/{account}/insights?' + urllib.parse.urlencode(params)
        while url:
            with urllib.request.urlopen(url) as r:
                j = json.loads(r.read().decode('utf-8'))
            for x in j.get('data', []):
                names[x['ad_id']] = x.get('ad_name', '')
                names[x['adset_id']] = x.get('adset_name', '')
                names[x['campaign_id']] = x.get('campaign_name', '')
                spend.append([x['date_start'], x['campaign_id'], x['adset_id'], x['ad_id'],
                              round(float(x.get('spend') or 0)), int(x.get('impressions') or 0), int(x.get('clicks') or 0)])
            url = j.get('paging', {}).get('next')
            time.sleep(0.2)
    return {'names': names, 'spend': spend}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dias', type=int, default=90)
    args = ap.parse_args()
    hoy = datetime.now(AR_TZ).date()
    desde = (hoy - timedelta(days=args.dias - 1)).isoformat()
    hasta = hoy.isoformat()

    ga_daily, ga_ads, ga_users = fetch_ga(desde, hasta)
    orders = fetch_orders(desde, hasta)
    meta = None
    try:
        meta = fetch_meta(desde, hasta)
    except Exception as e:  # la pauta es opcional; no frenar el reporte
        print('Meta API no disponible:', e)
    gsc = None
    try:
        gsc = fetch_gsc(desde, hasta)
    except Exception as e:  # Search Console también es opcional
        print('Search Console no disponible:', e)

    names_file = HERE / 'meta_ad_names.json'
    names = json.loads(names_file.read_text(encoding='utf-8')) if names_file.exists() else {}
    if meta:
        names.update({k: v for k, v in meta['names'].items() if v})

    data = {
        'generated': datetime.now(AR_TZ).strftime('%Y-%m-%d %H:%M'),
        'desde': desde, 'hasta': hasta, 'fix_date': FIX_DATE,
        'channels': [{'key': k, 'label': l} for k, l, _ in CHANNELS],
        'ga_daily': ga_daily, 'ga_ads': ga_ads, 'ga_users': ga_users,
        'orders': orders, 'names': names,
        'meta_spend': meta['spend'] if meta else None,
        'gsc_daily': gsc[0] if gsc else None, 'gsc_pages': gsc[1] if gsc else None,
        'gsc_hasta': gsc[2] if gsc else None, 'brand': list(BRAND),
    }
    tpl = (HERE / 'marketing_template.html').read_text(encoding='utf-8')
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    out = HERE / 'Marketing_Printy.html'
    out.write_text(tpl.replace('__MARKETING_DATA_JSON__', payload), encoding='utf-8')
    print(f'OK {out} | {desde}..{hasta} | GA filas {len(ga_daily)} / anuncios {len(ga_ads)} | pedidos {len(orders)}'
          f" | Meta {'sí' if meta else 'no'} | Search Console {'sí' if gsc else 'no'}")


if __name__ == '__main__':
    main()
