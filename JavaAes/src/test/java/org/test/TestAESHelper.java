package org.test;

import java.io.UnsupportedEncodingException;

import org.apache.commons.codec.binary.Hex;
import org.junit.Assert;
import org.junit.Test;

public class TestAESHelper {

	@Test
	public void test() {
		try {
			String key = "hawkeyes";
			String text = "中国人";

			byte[] enc = AESHelper.AESEncrypt(text.getBytes("UTF-8"), key.getBytes("UTF-8"));
			byte[] dec = AESHelper.AESDecrypt(enc, key.getBytes("UTF-8"));

			Assert.assertArrayEquals(text.getBytes("UTF-8"), dec);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Test
	public void test3() throws Exception {
		String key = "hawkeyes";

		// vb计算的加密结果
		String[] data = { "Text1", "ySr6kxYv3N/5GqzkKSLygQ==", "中国", "Ljt713+xq6rj2wdFpNRwiQ==", "Text1中国",
				"ivRuZ07mtWamAXostX0MLQ==", "emsno=abc1234567", "OqbQFj+ICmOj+jxQgvcZPg==" };

		for (int i = 0; i < data.length; i = i + 2) {
			Assert.assertEquals("enc", data[i + 1], AESHelper.Utf8AESBase64Encrypt(data[i], key));
			Assert.assertEquals("dec", data[i], AESHelper.Utf8AESBase64Decrypt(data[i + 1], key));
		}
	}

	@Test
	public void test4() throws Exception {
		String key = "hawkeyes333";

		// vb计算的加密结果
		String[] data = { "Text1", "y7iGOEbV7cu6oJN3Stzvfw==", "中国", "gO9vtX/5ej9XwoONx6IC+A==", "Text1中国",
				"PBnAlQsFABOj9AQzfwMtaA==", "emsno=abc1234567", "kguw8YJpKE7MxSOrF+qu7Q==" };

		for (int i = 0; i < data.length; i = i + 2) {
			Assert.assertEquals("enc", data[i + 1], AESHelper.Utf8AESBase64Encrypt(data[i], key));
			Assert.assertEquals("dec", data[i], AESHelper.Utf8AESBase64Decrypt(data[i + 1], key));
		}
	}

	@Test
	public void test2() {
		String tmp = "hello world";
		byte[] t = null;
		try {
			t = tmp.getBytes("UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}

		Assert.assertEquals("1", Hex.encodeHexString(t).toUpperCase(), bytes2hex(t));
	}

	private static String bytes2hex(byte[] bytes) {
		StringBuilder sb = new StringBuilder();
		for (int i = 0; i < bytes.length; i++) {
			String temp = (Integer.toHexString(bytes[i] & 0XFF));
			if (temp.length() == 1) {
				temp = "0" + temp;
			}
			sb.append(temp);
			sb.append("");
		}
		return sb.toString().toUpperCase();
	}
}
