# AESVbJava

vb和java能互相加解密的AES算法

使用方法：
* vb加密，java解密

  1. vb引用modEncode.bas、AES.cls两个文件
  2. vb系统内拼参数字符串，如`"entryid=530120161123456789"`
  3. 调用encodeUrlParams函数，加密参数字符串（参数为待加密串和密码字符串）,`encodeUrlParams(txtSource, "abc12345")`， `"abc12345"`为密码。
  4. 拼接完整的地址，如：`http://10.99.1.200/query.jsp?at70Pzk4zgJe4KMhtTZ7wt3EpV7T5M2n1tbSLjs%2BYT4%3D`
  5. 访问如上地址即可

  5. java，引用AESHelper类
  6. 得到加密的部分params，本例中为：`at70Pzk4zgJe4KMhtTZ7wt3EpV7T5M2n1tbSLjs%2BYT4%3D`
  8. 调用函数对加密部分解密
  `AESHelper.Utf8AESBase64Decrypt(URLDecoder.decode(params, "UTF-8"), "abc12345")`得到原始字符串：`"entryid=530120161123456789"`。`"abc12345"`为密码。
  9. 得到完整的访问地址：`http://10.99.1.200/query.jsp?entryid=530120161123456789`，服务端转向访问即可。

* java加密，java解密

  1. 加入AESHelper类
  1. 拼接参数字符串，如`"entryid=530120161123456789"`
  3. 调用`AESHelper.Utf8AESBase64Encrypt(txtSource, "abc12345")`进行加密
  4. 为了能让Base64字符串在url里进行传输，调用`URLEncode.encode(params, "UTF-8")`进行参数url编码
  5. 解密部分，同上