package org.example;

import lombok.Data;

@Data
public class Image {
    byte[] bytes;

    public Image(byte[] bytes) {
        this.bytes = bytes;
    }
}
