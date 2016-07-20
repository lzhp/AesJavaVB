package org.test;

import java.util.Arrays;

import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;

import org.apache.commons.codec.binary.Base64;

/**
 * 
 * @author lizhipeng
 *
 */
public class AESHelper {

	/**
	 * 对一个字符串进行AES加密，返回值Base64编码，字符串使用utf-8编码
	 * 
	 * @param plainText
	 * @param key
	 * @return
	 * @throws Exception
	 */
	public static String Utf8AESBase64Encrypt(String plainText, String key) throws Exception {
		byte[] enc = AESEncrypt(plainText.getBytes("UTF-8"), key.getBytes("UTF-8"));
		return Base64.encodeBase64String(enc);
	}

	/**
	 * AES解密， 传入base64编码的加密串，返回解密后的字符串
	 * 
	 * @param cipherText
	 * @param key
	 * @return
	 * @throws Exception
	 */
	public static String Utf8AESBase64Decrypt(String cipherText, String key) throws Exception {
		byte[] enc = AESDecrypt(Base64.decodeBase64(cipherText), key.getBytes("UTF-8"));
		return new String(enc, "UTF-8");
	}

	public static byte[] AESEncrypt(byte[] text, byte[] key) throws Exception {

		text = zeroPadding(text);
		key = zeroPadding(key, 16);

		SecretKeySpec aesKey = new SecretKeySpec(key, "AES");
		Cipher cipher = Cipher.getInstance("AES/ECB/NoPadding");
		cipher.init(Cipher.ENCRYPT_MODE, aesKey);
		return cipher.doFinal(text);
	}

	public static byte[] AESDecrypt(byte[] text, byte[] key) throws Exception {

		key = zeroPadding(key, 16);

		SecretKeySpec aesKey = new SecretKeySpec(key, "AES");
		Cipher cipher = Cipher.getInstance("AES/ECB/NoPadding");
		cipher.init(Cipher.DECRYPT_MODE, aesKey);
		return trimZeros(cipher.doFinal(text));
	}

	private static byte[] zeroPadding(byte[] text) {
		int count = text.length / 16;
		if (text.length % 16 != 0) {
			count++;
		}
		text = Arrays.copyOf(text, count * 16);
		return text;
	}

	private static byte[] zeroPadding(byte[] text, int length) {
		text = Arrays.copyOf(text, length);
		return text;
	}

	private static byte[] trimZeros(byte[] text) {
		int len = 0;
		for (int i = text.length - 1; i > 0; i--) {
			if (text[i] == 0) {
				len++;
			} else {
				break;
			}
		}
		if (len != 0) {
			text = Arrays.copyOf(text, text.length - len);
		}
		return text;
	}
}
