# springmvc-pageoffice

### 一、简介

 springmvc-pageoffice项目演示了在SpringMVC框架下如何使用PageOffice产品，此项目演示了PageOffice产品近90%的功能，每个PageOffice功能模块都以一个单独的Controller方式进行了展示，方便初学者学习和使用PageOffice产品。

### 二、项目环境要求

 Intelij IDEA,jdk1.8及以上版本

### 三、项目运行步骤

1. 使用git clone或者直接下载项目压缩包到本地并解压缩。
2. 导入springmvc-pageoffice项目到idea。
3. 配置tomcat。
4. 运行项目：点击运行按钮即可。
5. 启动浏览器访问：http://localhost:8080/ ，即可在线打开、编辑、保存office文件 。

### 五、PageOffice序列号

 PageOfficeV5.0标准版试用序列号：I2BFU-MQ89-M4ZZ-ZWY7K
 
 PageOfficeV5.0专业版试用序列号：DJMTF-HYK4-BDQ3-2MBUC

### 六、集成PageOffice到您的项目中的关键步骤

1. 在您项目的pom.xml中通过下面的代码引入PageOffice依赖。

   > pageoffice.jar的所有Releases版本都已上传到了Maven中央仓库，具体您要引用哪个版本请在Maven中央仓库地址中查看，建议使用Maven中央仓库中pageoffice.jar的最新版本。(Maven中央仓库中pageoffice.jar的地址：https://mvnrepository.com/artifact/com.zhuozhengsoft/pageoffice)

```
<dependency>
     <groupId>com.zhuozhengsoft</groupId>   
  <artifactId>pageoffice</artifactId>   
  <version>5.3.0.3</version>
</dependency>
```

1. 在您项目的web.xml配置如下代码。

```

    <servlet>
        <servlet-name>poserver</servlet-name>
        <servlet-class>com.zhuozhengsoft.pageoffice.poserver.Server</servlet-class>
    </servlet>
    <servlet-mapping>
        <servlet-name>poserver</servlet-name>
        <url-pattern>/poserver.zz</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>poserver</servlet-name>
        <url-pattern>/sealsetup.exe</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>poserver</servlet-name>
        <url-pattern>/posetup.exe</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>poserver</servlet-name>
        <url-pattern>/pageoffice.js</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>poserver</servlet-name>
        <url-pattern>/jquery.min.js</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>poserver</servlet-name>
        <url-pattern>/pobstyle.css</url-pattern>
    </servlet-mapping>
```

3.新建Controller并调用PageOffice，例如：

```
public class PageOfficeController {
@RequestMapping(value = "/Word", method = RequestMethod.GET)
  public ModelAndView showWord(HttpServletRequest request) {
  PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
  poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
  poCtrl.webOpen("/doc/test.doc", OpenModeType.docNormalEdit, "张三");
  request.setAttribute("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
  ModelAndView mv = new ModelAndView("Word");
   return mv;
  }
}
```

4.新建View页面,例如：Word.jsp（PageOfficeCtroller返回的View页面，用来嵌入PageOffice控件)，PageOffice在View页面输出的代码如下：

```
<div style=" width:auto; height:700px;">
    ${pageoffice}
</div>
```

### 七、电子印章功能说明

 如果您的项目要用到PageOffice自带电子印章功能，请按下面的步骤进行操作。

 1.请在您的web项目的web.xml中加上印章功能相关配置代码，具体代码请参考当前项目springmvc-pageoffice的web.xml中电子印章功能配置代码。

```
 <servlet>
        <servlet-name>adminseal</servlet-name>
        <servlet-class>com.zhuozhengsoft.pageoffice.poserver.AdminSeal</servlet-class>
    </servlet>
    <servlet-mapping>
        <servlet-name>adminseal</servlet-name>
        <url-pattern>/adminseal.zz</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>adminseal</servlet-name>
        <url-pattern>/loginseal.zz</url-pattern>
    </servlet-mapping>
    <servlet-mapping>
        <servlet-name>adminseal</servlet-name>
        <url-pattern>/sealimage.zz</url-pattern>
    </servlet-mapping>
    <mime-mapping>
        <extension>mht</extension>
        <mime-type>message/rfc822</mime-type>
    </mime-mapping>
    <context-param>
        <param-name>adminseal-password</param-name>
        <param-value>111111</param-value>
    </context-param>
```

 2.请拷贝当前项目根目录下poseal.db文件到您的web项目的WEB-INF/lib文件夹下。

 3.请参考当前项目的[一、15、演示加盖印章和签字功能（以Word为例）] 示例代码进行盖章功能调用。

### 八、 PageOffice开发帮助

 1.[Java API文档](https://www.zhuozhengsoft.com/help/java3/index.html)

 2.[JS API文档](https://www.zhuozhengsoft.com/help/js3/index.html)

 3.[PageOffice从入门到精通](https://www.kancloud.cn/pageoffice_course_group/pageoffice_course/646953)

 技术支持：https://www.zhuozhengsoft.com/Technical/

### 九、联系我们

 卓正官网：[https://www.zhuozhengsoft.com](https://www.zhuozhengsoft.com/)

 联系电话：400-6600-770

QQ: 800038353
