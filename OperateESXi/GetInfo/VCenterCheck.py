# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from __future__ import print_function
import ssl
import urllib
import urllib2
import cookielib
import threading
import time

from django.http import HttpResponse
from GetInfo.models import VCenterInfo, VCenterAliveStatus


def update_vcenter_alive_status(vcenter_ip):
    try:
        vc = VCenterInfo.objects.get(IP=vcenter_ip)
    except Exception:
        return 'VCenterInfo does not match query where IP is %s' % vcenter_ip
    else:
        url = vc.CheckUrl
        authorization = vc.Authorization
        status_id = vc.Status_id
        name = vc.Status.Name

    ssl_temp = ssl._create_default_https_context
    ssl._create_default_https_context = ssl._create_unverified_context

    headers = {
        'Authorization': authorization
    }

    cookieJar = cookielib.CookieJar()
    opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookieJar))
    req = urllib2.Request(url, headers=headers)
    result = opener.open(req)
    ssl._create_default_https_context = ssl_temp

    now_status = 'Alive' if result else 'Down'
    vc_state = VCenterAliveStatus.objects.get(Name=name)
    vc_state.LastUpdateStatus = now_status
    vc_state.save()


def create_vcenter_alive_info(ip, authorization, url, name):
    VCenterAliveStatus.objects.get_or_create(Name=name)
    vc_state = VCenterAliveStatus.objects.get(Name=name)
    state_id = vc_state.id
    VCenterInfo.objects.get_or_create(IP=ip, Status_id=state_id)
    vc_info = VCenterInfo.objects.filter(IP=ip)
    vc_info.update(Authorization=authorization, CheckUrl=url)

    update_vcenter_alive_status(ip)


def control_monitor(enable):
    with threading.Lock():
        count = 0
        while enable:
            count += 1
            for vc in list(VCenterInfo.objects.all().values()):
                vc_ip = vc['IP']
                update_vcenter_alive_status(vc_ip)
            print ('round %s' % count)
            time.sleep(60)
    if enable:
        # return HttpResponse('已启动检测 vcenter alive status control monitor')
        return 'started...'
    else:
        # return HttpResponse('已停止检测 vcenter alive status control monitor')
        return 'stoped...'
