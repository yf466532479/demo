package com.file.upordownfile;

import lombok.Data;
import lombok.experimental.Accessors;

import java.io.Serializable;

@Data
@Accessors(chain = true)
public class User implements Serializable {

    private String name;

    private String sex;

    private String age;

    private String birthday;



}
