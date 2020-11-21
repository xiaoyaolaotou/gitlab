import time
import gitlab
import pandas as pd
from io import BytesIO
from flask import Flask,render_template,request,make_response,send_file

app = Flask(__name__,template_folder="templates",static_folder="static")
app.debug=True

gl = gitlab.Gitlab('http://code.gome.inc/gitlab',private_token='ve9ifEu9ZtAeqWik1L_y',timeout=50, api_version='4')


# start_time = '2020-06-10T00:00:00Z'
# end_time = '2020-06-30T24:00:00Z'

    
@app.route("/gitlabv1", methods=['GET','POST'])
def gitlabv1():
    if request.method=='GET':
        return render_template("gitlab.html")
    else:
        t1 = request.form.get('t1')
        t2 = request.form.get('t2')
        list2 = []
        projects=gl.projects.list(owned=True)
        for project in projects:
            for branch in project.branches.list():
                commits=project.commits.list(all=True,query_parameters={'since': t1,'until':t2, 'ref_name': branch.name})
                for commit in commits:
                    com = project.commits.get(commit.id)
                    pro={}    
                    try:
                        #print(project.path_with_namespace,com.author_name,com.stats["total"])
                        pro["projectName"]=project.path_with_namespace
                        pro["authorName"]=com.author_name
                        pro["branch"]=branch.name
                        pro["additions"]=com.stats["additions"]
                        pro["deletions"]=com.stats["deletions"]
                        pro["commitNum"]=com.stats["total"]                  
                        list2.append(pro)
                    except :
                        print("有错误, 请检查")
        # 去重并计算
        ret = {}
        for ele in list2:
            key = ele["projectName"]+ele["authorName"]+ele["branch"]
            if key not in ret:
                ret[key] = ele
                ret[key]["commitTotal"] = 1
            else:
                ret[key]["additions"] += ele["additions"]
                ret[key]["deletions"] += ele["deletions"]
                ret[key]["commitNum"] += ele["commitNum"]
                ret[key]["commitTotal"] +=1

        list1 = []
        for key,v in ret.items():
            v["项目名"] = v.pop("projectName")
            v["开发者"] = v.pop("authorName")
            v["分支"] = v.pop("branch")
            v["添加代码行数"] = v.pop("additions")
            v["删除代码行数"] = v.pop("deletions")
            v["提交总行数"] = v.pop("commitNum")
            v["提交次数"] = v["commitTotal"]
            list1.append(v)
        print(list1)
        """
            导出excel
        """
        out = BytesIO()
        writer = pd.ExcelWriter(out, engine='xlsxwriter')
        data=pd.DataFrame(list1,columns=["项目名","开发者","分支","添加代码行数","删除代码行数","提交总行数","提交次数"])
        writer.save()
        writer.close()
        file_name = time.strftime('%Y%m%d', time.localtime(time.time())) + '.xlsx'
        response = make_response(out.getvalue())
        response.headers["Content-Disposition"] = "attachment; filename=%s" % file_name
        response.headers["Content-type"] = "application/x-xls"
        return render_template("gitlab.html",msg=list1), response
       


if __name__ == "__main__":
    app.run(port="1234")