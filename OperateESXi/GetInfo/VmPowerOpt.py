# -*- coding: utf-8 -*-
from GetInfo.models import VmInfo, ClusterInfo
from pyvim import connect
import time


def shutdown_vms(cluster_ip):
    # shut down the vms(all the vms in the cluster)
    # reset_vms_info(cluster_ip)
    shutdown_list = list(VmInfo.objects.filter(Cluster=cluster_ip, powerStatus='poweredOn').values())
    t_host = ClusterInfo.objects.get(IP=cluster_ip)
    user = t_host.User
    pwd = t_host.Pwd
    mysi = connect.ConnectNoSSL(host=cluster_ip, port=443, user=user, pwd=pwd)
    mysi_content = mysi.content
    vm_list = []

    for i in shutdown_list:
        vm = mysi_content.searchIndex.FindByUuid(None, i['instanceUuid'], True, True)
        vm_list.append(vm)

    count = 0
    to_shutdown = vm_list[:]
    done_list = []

    while count < 3:
        if len(to_shutdown) > 0:
            for i in to_shutdown:
                try:
                    i.ShutDownGuest()
                except Exception:
                    i.PowerOff()
            time.sleep(10)
            for i in to_shutdown[:]:
                if i.summary.runtime.powerState == 'poweredOff':
                    to_shutdown.remove(i)
                    done_list.append(i)
        else:
            break
        count += 1

    if len(to_shutdown) == 0:
        connect.Disconnect(mysi)
        return {'PowerOffFinish': 'All vms PoweredOff'}
    elif len(done_list) == 0:
        connect.Disconnect(mysi)
        return {'PowerOffFailed': 'No vm PoweredOff'}
    else:
        fail_list = []
        for i in to_shutdown:
            fail_list.append(i.summary.config.name)
        connect.Disconnect(mysi)
        return {'Following Vms PowerOff Failed': fail_list}


def start_up_vms(cluster_ip):
    power_on_list = list(VmInfo.objects.filter(Cluster=cluster_ip, powerStatus='poweredOn').values())
    t_host = ClusterInfo.objects.get(IP=cluster_ip)
    user = t_host.User
    pwd = t_host.Pwd
    mysi = connect.ConnectNoSSL(host=cluster_ip, port=443, user=user, pwd=pwd)
    mysi_content = mysi.content
    vm_list = []

    for i in power_on_list:
        vm = mysi_content.searchIndex.FindByUuid(None, i['instanceUuid'], True, True)
        vm_list.append(vm)

    count = 0
    to_start = vm_list[:]
    done_list = []

    while count < 3:
        if len(to_start) > 0:
            for i in to_start:
                i.PowerOn()
            time.sleep(10)
            for i in to_start[:]:
                if i.summary.runtime.powerState == 'poweredOn':
                    to_start.remove(i)
                    done_list.append(i)
        else:
            break
        count += 1

    if len(to_start) == 0:
        connect.Disconnect(mysi)
        return {'PowerOnFinish': 'All vms PoweredOn'}
    elif len(done_list) == 0:
        connect.Disconnect(mysi)
        return {'PowerOnFailed': 'No vm PoweredOn'}
    else:
        for i in to_start:
            myname = i.summary.config.name
            to_start[to_start.index(i)] = myname
        connect.Disconnect(mysi)
        return {'Following Vms PowerOn Failed': to_start}
