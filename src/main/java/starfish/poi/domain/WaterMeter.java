package starfish.poi.domain;

import lombok.Data;

@Data
public class WaterMeter {

    private Integer id;

    private String name;

    private String time;

    private Float data;

    private String status;

}
