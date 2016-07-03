#-*-encoding:utf8-*-
import xlrd, os, sys, re, time, platform

def open_xls(filepath):
    try:
        data = xlrd.open_workbook(filepath)
        return data
    except Exception, e:
        print e

#patterns
pt_bcm = re.compile('(//|/\*)?[^\n]*certificateServiceClient\.(sign|signWithInfo|verify|encrypt|decrypt|decode|encode|getMerCert|getValidCert|getValidCertInfo)\(\"(\w+)\"')
pt_ch = re.compile('\.setChannelSystemId\(\"(\w+)\"\)|\s*\"channelApi\",\"(\w+)\"')
pt_tr = re.compile('\.setTransTypeId\(\"(\w+)\"\);')

def biz(xls, txt):
    '''
    0.容器: 
        res_api_map: { channelApi + ',' + transTypeId : { bcm_code : set([ optype1, optype2 ])}}
        res_csi_map: { channelSystemid + ',' +transTypeId: { bcm_code : set([ optype1, optype2 ])} }
        bcm_grvid_map: { bcmcode : [grvid, grvid]}
        bcmcodes_unmapped : [bcmcode, bcmcode, ...] 从bcmcodes_unmapped.txt读取，没有对应渠道的bcmcodes
    1.加载excel表格
    2.遍历所有groovy_content内容并正则提取所有pt_bcm匹配到的结果放入bcm_grvid_map中
    3.读取bcmcodes_unmapped.txt读取所有没有对应渠道的bcmcode，并放入bcmcodes_unmapped
        3.1正则匹配bcm_code, channelSystemId/channelApi, transTypeId, 将匹配到的bcm_code和它所在行放入bcm_grvid_map
    4.遍历bcmcodes_unmapped元素，从bcm_grv_map中取出对应groovy所在行列表
        4.1根据行列表取出groovy_content内容
        4.2正则匹配channelSystemId or channelApi , transTypeId, optype元素
        4.3根据是channelSystemId还是channelApi放入不同的容器中
    5.log输出
    '''
    #step 0
    global pt_bcm, pt_ch, pt_tr
    res_api_map = {}
    res_csi_map = {}
    bcm_grvid_map = {}
    bcmcodes_unmapped = []

    #step 1
    workbook = open_xls(xls)
    sheet = workbook.sheets()[0]
    nrows = sheet.nrows
    ncols = sheet.ncols

    #step 2
    for rindex in range(2,nrows):
        #groovy脚本在第五列
        groovy_content = str(sheet.cell(rindex, 4))
        rs = pt_bcm.search(groovy_content)
        if rs is not None:
            valid = True if rs.group(1) is None else False
            optype = rs.group(2).strip()
            code = rs.group(3).strip()
            #valid表示certificateServiceClient是否在注释内
            if valid:
                if bcm_grvid_map.has_key(code):
                    bcm_grvid_map[code].append(rindex)
                else:
                    bcm_grvid_map[code] = [rindex]
            else:
                print 'annotationed bcmcode:',code
    print 'step2 finished. bcm_grvid_map:'
    #print bcm_grvid_map
    cnt = 0
    keylist = sorted(bcm_grvid_map.items(), key = lambda a: a[0])
    for k,v in keylist:
        cnt += 1
        print '{}  {} : {}'.format(cnt, k, v)
    print '=='*20
    return

    #step 3
    with open(txt) as f:
        for line in f:
            #日志格式：1, xxx
            bcmcodes_unmapped.append(line.split(',')[1].strip())
    print 'step 3 finished. bcmcodes_unmapped:'
    print bcmcodes_unmapped
    print len(bcmcodes_unmapped)
    print '=='*20

    #step 4
    count = 0
    #遍历bcmcodes_unmapped
    for umbcm in bcmcodes_unmapped: 
        if bcm_grvid_map.has_key(umbcm):
            #step 4.1
            pt_bcm = re.compile('(//|/\*)?[^\n]*certificateServiceClient\.(sign|signWithInfo|verify|encrypt|decrypt|decode|encode|getMerCert|getValidCert|getValidCertInfo)\(\"({})\"'.format(umbcm))
            groovy_content_rid_list = bcm_grvid_map[umbcm]
            for rid in groovy_content_rid_list:
                grv = str(sheet.cell(rid,4))
                rs_bcm = pt_bcm.search(grv)
                rs_ch = pt_ch.search(grv)
                rs_tr = pt_tr.search(grv)
                if rs_bcm is not None and rs_ch is not None and rs_tr is not None:
                    #这里不考虑注释的情况，因为在初始化bcm_grvid_map时已经筛选过
                    channelSystemId = rs_ch.group(1)
                    channelApi = rs_ch.group(2)
                    transTypeId = rs_tr.group(1)
                    bcm_code = rs_bcm.group(3)
                    optype = rs_bcm.group(2)
                    #首先判断是否有channelApi或channelSystemId
                    if channelSystemId is not None and channelApi is None:
                        key = channelSystemId + ',' + transTypeId
                        if res_csi_map.has_key(key):
                            if res_csi_map[key].has_key(umbcm):
                                res_csi_map[key][umbcm].add(optype)
                            else:
                                opset = set()
                                opset.add(optype)
                                res_csi_map[key][umbcm] = opset
                    elif channelApi is not None and channelSystemId is None:
                        key = channelApi + ',' + transTypeId
                        if res_api_map.has_key(key):
                            if res_api_map[key].has_key(umbcm):
                                res_api_map[key][umbcm].add(optype)
                            else:
                                opset = set()
                                opset.add(optype)
                                res_api_map[key][umbcm] = opset
                    else:
                        #没有channelSystemId也没有channelApi
                        print '!!!!!!no channelApi or channelSystemId:umbcm:{}'.format(umbcm)
                        continue
                else:
                    print 'rs_bcm, rs_ch, rs_tr:{},{},{}'.format(rs_bcm, rs_ch, rs_tr  )

        else:
            count += 1
            print "{}th this bcm code is not contained in bcm_grvid_map:{}".format(count, umbcm)
    print 'step 4 finished.'
    print '=='*20

    #step 5
    print 'res_api_map:'
    print res_api_map
    print 'res_csi_map:'
    print res_csi_map

if __name__ == '__main__':
    sys_type = platform.system()
    xls = ''
    txt = ''
    if sys_type == 'Windows':
        xls = 'groovy_contents.xls'
        txt = 'data\\bcmcode_unmapped.txt'
    else:
        xls = './groovy_content.xls'
        txt = './data/bcmcode_unmapped.txt'
    start = time.time()
    biz(xls, txt)
    end = time.time()
    print 'time consume:{}s'.format(end-start)