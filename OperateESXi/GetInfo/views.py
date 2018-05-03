# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from __future__ import print_function

from django.shortcuts import render
from django.http import HttpResponse, HttpResponseBadRequest
from django.views.decorators.csrf import csrf_exempt

from GetInfo.models import ClusterInfo, OldClusterInfo
from GetInfo.forms import ClusterForm, TypeInClusterForm

from GetInfo.VmInfoOpt import get_vms_info, reset_vms_info, create_vms_info
from GetInfo.VmStatusOpt import get_vms_status
from GetInfo.VmPowerOpt import shutdown_vms, start_up_vms
from GetInfo.AutoStartConfig import auto_start
from VCenterCheck import control_monitor

from pyvim import connect
import datadiff
import json
import os
# Create your views here.


def check_request(body):
    try:
        received_json_data = json.loads(body)
    except ValueError:
        # return 'ValueError: No JSON Object could be decoded - from check_request'
        return {'ValueError: No JSON Object could be decoded - from check_request'}
    try:
        cluster_ip = received_json_data['cluster_ip']
    except KeyError:
        # return 'KeyError: Content Keyword cluster_ip Not Found - from check_request'
        return {'KeyError: Content Keyword cluster_ip Not Found - from check_request'}
    return cluster_ip


def check_return(mydata, ori):
    if isinstance(mydata, list):
        return HttpResponse(json.dumps(mydata))
    elif isinstance(mydata, dict):
        return HttpResponse(json.dumps(mydata), status=400)
    else:
        return HttpResponse('rua! rua! rua! - from %s' % ori, status=400)

#######################################################################
# The following definitions are used to return data to the front side #
#######################################################################
@csrf_exempt
def index(request):
    if request.method == 'POST':
        cluster_form = ClusterForm(request.POST)
        if cluster_form.is_valid():
            cluster_ip = cluster_form.cleaned_data["Cluster"]
            show_list = []
            if 'getstatus' in request.POST:
                show_list = get_vms_status(cluster_ip)
            elif 'getinfo' in request.POST:
                show_list = get_vms_info(cluster_ip)
            elif 'resetinfo' in request.POST:
                show_list = reset_vms_info(cluster_ip)
            elif 'shutdown' in request.POST:
                show_list = shutdown_vms(cluster_ip)
            elif 'poweron' in request.POST:
                show_list = start_up_vms(cluster_ip)
            elif 'createvmsinfo' in request.POST:
                show_list = create_vms_info(cluster_ip)
            # return HttpResponse(json.dumps(show_list))
            return HttpResponse(show_list)
    else:
        cluster_form = ClusterForm()
    return render(request, 'index.html', {'cluster_form': cluster_form})


def test(request):
    return render(request, 'getinfo/mytest.html')


@csrf_exempt
def upload(request):
    if request.method == 'POST':
        obj = request.FILES.get('myfile')
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        f = open(os.path.join(BASE_DIR, 'static', 'pic', obj.name), 'a+')
        for chunk in obj.chunks():
            f.write(chunk)
        f.close()
        return HttpResponse(obj.name + 'OK')


@csrf_exempt
def body_back(request):
    # try:
    #     received_json_data = json.dumps(request.body)
    # except Exception:
    #     return HttpResponse('ValueError')
    # return HttpResponse(json.loads(received_json_data))
    return HttpResponse(json.dumps({"PowerOffFinish": "All vms PoweredOff"}))


@csrf_exempt
def uri_back(request):
    received_data = request.GET.get('check_vc_state')
    return HttpResponse(received_data)


@csrf_exempt
def api(request):
    if request.method == 'POST':
        try:
            received_json_data = json.loads(request.body)
        except Exception:
            return HttpResponse('ValueError')
        try:
            cluster_ip = received_json_data['cluster']
        except Exception:
            return HttpResponse('clusterip Error')
        try:
            method = received_json_data['method']
        except Exception:
            return HttpResponse('method Error')
        if method == 'getstatus':
            show_list = get_vms_status(cluster_ip)
        elif method == 'getinfo':
            show_list = get_vms_info(cluster_ip)
        elif method == 'resetinfo':
            show_list = reset_vms_info(cluster_ip)
        elif method == 'shutdown':
            show_list = shutdown_vms(cluster_ip)
        elif method == 'poweron':
            show_list = start_up_vms(cluster_ip)
        elif method == 'createvmsinfo':
            show_list = create_vms_info(cluster_ip)
        else:
            return HttpResponse('ERROR * 2!')
        return HttpResponse(json.dumps(show_list))


@csrf_exempt
def vms_status(request):
    if request.method == 'GET':
        cluster_ip = request.GET.get('cluster_ip')
        if cluster_ip == '' or cluster_ip is None:
            return HttpResponse('Cluster IP import Error, Please Check - from vms_status.GET', status=400)
        else:
            show_list = get_vms_status(cluster_ip)
            if isinstance(show_list, list):
                return HttpResponse(json.dumps(show_list))
            elif isinstance(show_list, dict) and len(show_list.values()) == 1 and 'error' in show_list.keys():
                return HttpResponse(json.dumps(show_list), status=400)
            else:
                return HttpResponse('rua! rua! rua! - from vms_status.GET')
    else:
        return HttpResponse('Error : Can Only Use GET Request - from vms_status.GET', status=400)


@csrf_exempt
def vms_info(request):
    if request.method == 'GET':
        cluster_ip = request.GET.get('cluster_ip')
        if cluster_ip == '' or cluster_ip is None:
            return HttpResponse('Cluster IP import Error, Please Check - from vms_info.GET', status=400)
        else:
            show_list = get_vms_info(cluster_ip)
        if isinstance(show_list, list):
            return HttpResponse(json.dumps(show_list))
        elif isinstance(show_list, dict):
            return HttpResponse(json.dumps(show_list), status=400)
        else:
            return HttpResponse('rua! rua! rua! - from vms_info.GET')

    elif request.method == 'POST':
        try:
            received_json_data = json.loads(request.body)
        except ValueError:
            return HttpResponse('ValueError: No JSON Object could be decoded - from vms_info', status=400)
        try:
            cluster_ip = received_json_data['cluster_ip']
        except KeyError:
            return HttpResponse('KeyError: Content Keyword cluster_ip Not Found - from vms_info', status=400)
        show_list = reset_vms_info(cluster_ip)
        if (len(show_list.values()) > 1) or (len(show_list.values()) == 1 and 'warning' in show_list.keys()):
            return HttpResponse(json.dumps(show_list))
        elif isinstance(show_list, dict) and len(show_list.values()) == 1 and 'error' in show_list.keys():
            return HttpResponse(json.dumps(show_list), status=400)
        else:
            return HttpResponse('rua! rua! rua! - from vms_info.POST')
    else:
        return HttpResponse('Error: Can Only Use GET, POST Request - from vms_info', status=400)


@csrf_exempt
def vms_power_off(request):
    if request.method == 'POST':
        try:
            received_json_data = json.loads(request.body)
        except ValueError:
            return HttpResponse('ValueError: No JSON Object could be decoded - from vms_power_off', status=400)
        try:
            cluster_ip = received_json_data['cluster_ip']
        except KeyError:
            return HttpResponse('KeyError: Content Keyword cluster_ip Not Found - from vms_power_off', status=400)
        show_list = shutdown_vms(cluster_ip)
        return HttpResponse(json.dumps(show_list))
    else:
        return HttpResponse('error: Can Only Use POST Request - from vms_power_off', status=400)


@csrf_exempt
def vms_power_on(request):
    if request.method == 'POST':
        try:
            received_json_data = json.loads(request.body)
        except ValueError:
            return HttpResponse('ValueError: No JSON Object could be decoded - from vms_power_on', status=400)
        try:
            cluster_ip = received_json_data['cluster_ip']
        except KeyError:
            return HttpResponse('KeyError: Content Keyword cluster_ip Not Found - from vms_power_on', status=400)
        show_list = start_up_vms(cluster_ip)
        return HttpResponse(json.dumps(show_list))
    else:
        return HttpResponse('error: Can Only Use POST Request - from vms_power_on', status=400)


@csrf_exempt
def vms_auto_start(request):
    if request.method == 'POST':
        try:
            received_json_data = json.loads(request.body)
        except ValueError:
            return HttpResponse('ValueError: No JSON Object could be decoded - from vms_power_on', status=400)
        try:
            cluster_ip = received_json_data['cluster_ip']
            start_opt = received_json_data['start_opt']
        except KeyError:
            return HttpResponse('KeyError: Content Keyword cluster_ip Not Found - from vms_power_on', status=400)
        if start_opt == 1:
            start_opt = True
        elif start_opt == 0:
            start_opt = False
        else:
            return {'error': 'Error with start_opt, the value is: %s - from vms_auto_start' % start_opt}
        show_list = auto_start(cluster_ip, start_opt)
        return HttpResponse(json.dumps(show_list))
    else:
        return HttpResponse('error: Can Only Use POST Request - from vms_auto_start', status=400)


@csrf_exempt
def type_in_cluster(request):
    if request.method == 'POST':
        typein_form = TypeInClusterForm(request.POST)
        if typein_form.is_valid():
            cluster_name = typein_form.cleaned_data['Name']
            cluster_ip = typein_form.cleaned_data['IP']
            cluster_user = typein_form.cleaned_data['User']
            cluster_pwd = typein_form.cleaned_data['Pwd']

            new_info_dict = {'Name': cluster_name, 'IP': cluster_ip, 'User': cluster_user, 'Pwd': cluster_pwd}

            t_info = ClusterInfo.objects.filter(IP=cluster_ip)

            if len(t_info) == 1:
                t_info_dict = list(t_info.values())[0]
                t_info_dict.pop('id')
                if cmp(new_info_dict, t_info_dict):
                    old_info = OldClusterInfo.objects.filter(IP=cluster_ip)
                    if len(old_info) == 1:
                        old_info_dict = list(old_info.values())[0]
                        old_info_dict.pop('id')
                        if cmp(t_info_dict, old_info_dict):
                            old_info.delete()
                    OldClusterInfo.objects.get_or_create(**t_info_dict)
                    t_info.update(Name=cluster_name, User=cluster_user, Pwd=cluster_pwd)
                    return HttpResponse(json.dumps({'now': new_info_dict,
                                                    'old': t_info_dict,
                                                    }))
                else:
                    return HttpResponse(json.dumps({'Warning': 'The information is same, Noting to do!'}))
            else:
                ClusterInfo.objects.get_or_create(**new_info_dict)
                return HttpResponse(json.dumps({'Create New Cluster Finish': 'Create New Cluster Finish!'}))
    else:
        typein_form = TypeInClusterForm()
    return render(request, 'typein.html', {'typein_form': typein_form})


@csrf_exempt
def multi_type_in_cluster(request):
    if request.method == 'POST':
        try:
            received_json_data = json.loads(request.body)
        except ValueError:
            return HttpResponse('ValueError: No JSON Object could be decoded')
        error_list = []
        done_list = []
        del_list = []
        for new_info_dict in received_json_data:
            if len(new_info_dict.keys()) == 4 and {'Name', 'IP', 'User', 'Pwd'}.issubset(new_info_dict):
                cluster_name = new_info_dict['Name']
                cluster_ip = new_info_dict['IP']
                user = new_info_dict['User']
                pwd = new_info_dict['Pwd']
                try:
                    mysi = connect.ConnectNoSSL(host=cluster_ip, port=443, user=user, pwd=pwd)
                except Exception:
                    error_list.append(new_info_dict)
                else:
                    connect.Disconnect(mysi)
                    t_info = ClusterInfo.objects.filter(IP=cluster_ip)
                    if len(t_info) == 1:
                        t_info_dict = list(t_info.values())[0]
                        t_info_dict.pop('id')
                        if datadiff.diff([t_info_dict], [new_info_dict]):
                            old_info = OldClusterInfo.objects.filter(IP=cluster_ip)
                            if len(old_info) == 1:
                                old_info_dict = list(old_info.values())[0]
                                old_info_dict.pop('id')
                                if datadiff.diff([old_info_dict], [t_info_dict]):
                                    old_info.delete()
                                    del_list.append(old_info_dict)
                            OldClusterInfo.objects.get_or_create(**t_info_dict)
                            t_info.update(Name=cluster_name, User=user, Pwd=pwd)
                        done_list.append(new_info_dict)
                    else:
                        ClusterInfo.objects.get_or_create(**new_info_dict)
                        done_list.append(new_info_dict)
            else:
                error_list.append(new_info_dict)
        return HttpResponse(json.dumps({'Done': done_list,
                                        'Error': error_list,
                                        'Delete': del_list
                                        }))
    else:
        return HttpResponse('Can Only Use POST Request - from multi_type_in_cluster')
