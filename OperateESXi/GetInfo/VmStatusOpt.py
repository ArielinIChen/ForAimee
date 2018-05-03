# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from __future__ import print_function
from GetInfo.models import ClusterInfo
from pyvim import connect


# Get vm status from esxi
def collect_vm_status(vm):
    summary = vm.summary
    my_status = dict()
    my_status['Name'] = summary.config.name
    my_status['Path'] = summary.config.vmPathName
    my_status['Guest'] = summary.config.guestFullName
    my_status['Annotation'] = summary.config.annotation
    my_status['instanceUuid'] = summary.config.instanceUuid
    my_status['powerStatus'] = summary.runtime.powerState
    ip = summary.guest.ipAddress

    if ip is None or ip == '':
        my_status['IP'] = 'unset'
    else:
        my_status['IP'] = ip
    return my_status


def get_vms_status(cluster_ip):
    # show vms(all the vms in the cluster) status from esxi by using function-collect_vm_status
    filter_hosts = ClusterInfo.objects.filter(IP=cluster_ip)

    if len(filter_hosts) == 1:
        t_host = ClusterInfo.objects.get(IP=cluster_ip)
        user = t_host.User
        pwd = t_host.Pwd
        mysi = connect.ConnectNoSSL(host=cluster_ip, port=443, user=user, pwd=pwd)
        mysi_content = mysi.content

        for child in mysi_content.rootFolder.childEntity:
            if hasattr(child, 'vmFolder'):
                datacenter = child
                vmFolder = datacenter.vmFolder
                vmList = vmFolder.childEntity
                status_list = []

                for vm in vmList:
                    # get vm status from esxi by function-collect_vm_status
                    status_dict = collect_vm_status(vm)
                    status_list.append(status_dict)
                connect.Disconnect(mysi)
                return status_list
            else:
                connect.Disconnect(mysi)
                return {'error': '%s hasattr vmFolder - from get_vms_status' % child}
    elif len(filter_hosts) > 1:
        return {'error': 'inner Error with: %s - from get_vms_status' % cluster_ip}
    else:
        return {'error': 'This IP: %s Not Found - from get_vms_status' % cluster_ip}
