<%@ page contentType="text/css;charset=UTF-8" %>
<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>

#p-logo {
	height: 86px;
	width: 107px;
	position: absolute;
	left: 50px;
	overflow: visible;
	z-index: 30;
	background: url('${c:encodeURL('/zssdemo/img/top_zk_logo.png')}') no-repeat transparent scroll left top;
}

#p-logo, #p-logo a, #p-logo a:hover {
	height: 86px;
    width: 107px;
    display: block;
    text-decoration: none;
}

#header {
	background: url("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAABbCAYAAABK4BTvAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAABMVJREFUeNqMV+2O4zYMlGVlsy36Sn3iPlRf4ICiwOGytvVRznAo+a5/kkWQrE2LnCE5ZLb8198jpS2l0VPKOfH7Zm+86pXSXviZeRHXHw83LnajV3/vuxkd9plT4ZMwaPa5m3GrfgqfNmflgyaFF69TxjLKdtJ5KIRBw5zqyaPNl8d4mUFvfmBesWe6O814s1OaGTye9j17GFkh2PWcjpc/fb4cCMIAgM1ctstdWwiFaHv3oHkD9+xmH/6JeA19pjvExNjwgFwCCLzB8PiyE4EUMSXRg+DxQtwAVyuvF94IQ9w4v/w7rsEDaWtmGG6ACECCbMQLtwBjTBSmCQZbgEhOEa7jf2Cw8DJdATku4AZOxLWuTD0+GEahW6QL9BDI5sZ48OPpjHTESP7qynVVbCovr66szMAN382N8EYN4BPomWugivp7/mb/wtXlscED7pu3TAB4CkgZQvL4mBXF32EYxAIp6WlOOGK/Lk+pxZnJF8gjb92Rwi1yj2pCrObGc4cTkVMWiH3//q/HBi8AZJVUUBnOlyocp37+7rQgYwwF9Yi4FDBRv14eBuN8eJzGqWemqPsQK9tWqSSvmQXiZRYU4eR4EYSEwKjLDJap0w3Qwnh3PxGHGKfeCjhtVs4ht92TIDCema6qiRBGWk2G/9mF0fhwdb3cKPLcTqF/ip5DffLx6aXFRivO6eFFXCbKHC2xeVzxAOI0oGU1/B1pnDicDbTruFSgIVAEUz0JUdAFqGGwiWC4RQ0i91TbLBE4zVA8UWtwClXiS2IgjWQ9hn4n9fDXjzSzhTBaVE+kie06VgrxXf2CMApPjK7rfWUEnEZnGgOeQqBttzSmW9Gq7/MUKPXvnDNoDXq6VBTIJ3oZ84MSmETLUNNtvO+SQt2uSyNDnKhkFzPjMYbKMhttcSndcaXgMJKoR+MHoFbnTCyuFE1ykmehul52nz9jhOvh6CLOmA5T29GuULHZAiq3JvljP206kU8N15kkevBmyaU5mnMQOiUvwPW24q1TALZoO59aWRMh5g3a3fVFU6rrJhXYTjzrrV1jbMQpJFouY+SxKKIDyWVb2hObgZLgG0AWmKzmT2OhphBcOjFGMWpyzkOB03DyzERO/ydUY86b4sWaFpesaolpbAIbMoP8XtdNQOvNqP8yPgCmax7iDSBDK83xQ+Oj1bUQhf7EaQHKvJQ5W0J/SH6WsuWpZnk+NctMRQyDycgWGu4VMiW53tE7z3kYiGFPDgt+GKiBv+enX+M9u77HqhCAYtng5M+r4SqKYqQl9mHAJpRiDG/jzPKnYp3OX97m4ubCVaUU476sjSX4IfbNM1V+qsEpBGNJi9SsrC1ZFR2v61jbsxmXuRQFgFirJ5jBCeZTITYn6HfU4mRhUz2iSoaQxuSv55JA7kVohWjR0X+unrHagEKyxu4tvtCiYIS7GRv95W72fEPqhJBwbvZR7gARud1VyBx/3vPZn9zWznNTMKfOB/+SlHDZY2I1FwNStN9+psQ42xRfum32vvewzsSZTuf+o2yJvewb/XU7UZvVriGQhsCA1P2mDGk11H3fKGvvGbexMRbAESnkDxttJvEDDTwGEK0LOP/P9MYLPHx71/Cfdwzh+o93Dcs7hv8JMACQz1xGCqi2mQAAAABJRU5ErkJggg==") repeat-x scroll left top transparent;
	height: 86px;
    margin: 0;
    min-height: 55px;
    position: relative;
    top: 0;
    width: 100%;
    z-index: 10;
}

#header-download {
	background: url('${c:encodeURL('/zssdemo/img/download-btn.png')}') no-repeat scroll 0 0 transparent;
	display: block;
	position: absolute;
	width: 178px;
    top: 29px;
	right: 50px;
	height: 38px;
}

#header-download:hover {
	background-position: 0 -38px;
} 