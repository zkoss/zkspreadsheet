/*
 * @(#)Clock.java
 *
 * Copyright (c) 1999-2016 7thOnline, Inc.
 * 24 W 40th Street, 11th Floor, New York, NY 10018, U.S.A.
 * All rights reserved.
 *
 * This software is the confidential and proprietary information of 7thOnline,
 * Inc. ("Confidential Information").  You shall not disclose such Confidential
 * Information and shall use it only in accordance with the terms of the
 * license agreement you entered into with 7thOnline.
 */
package zss.issue;

import java.util.Date;

public class Clock {
    public static final int TYPE_SECOND = 0;
    public static final int TYPE_MINUTE = 1;
    public static final int TYPE_HOUR   = 2;
    public static final int TYPE_DAY    = 3;
    
    private long startedTime;
    private long endedTime;
    
    public Clock() {
        
    }
    public Date getStartedTime() {
        return new Date(startedTime);
    }
    public Date getEndedTime() {
        return new Date(endedTime);
    }
    public void start() {
        startedTime = System.currentTimeMillis();
    }
    
    public void stop() {
        endedTime = System.currentTimeMillis();
    }
    
    public long elaspsedTime() {
        return endedTime - startedTime; 
    }
    
    public double elaspsedTime(int timeType) {
        return this.elaspsedTime(this.elaspsedTime(), timeType);
    }
    
    public String elaspsedTimeString() {
        long timeMillis = this.elaspsedTime();

        int days = (int)elaspsedTime(timeMillis, TYPE_DAY); 
        int hours = (int)elaspsedTime((timeMillis % (24 * 3600 * 1000.00)), TYPE_HOUR); 
        int minutes = (int)elaspsedTime((timeMillis % (3600 * 1000.00)), TYPE_MINUTE); 
        int seconds = (int)elaspsedTime((timeMillis % (60 * 1000.00)), TYPE_SECOND); 
        int millis = (int)(timeMillis % 1000.00);
        
        StringBuilder builder = new StringBuilder(100);
        if(days > 0){
            builder.append(days + "d ");
            builder.append(hours + "h ");
            builder.append(minutes + "m ");
            builder.append(seconds + "s ");
            builder.append(millis + "ms ");
        }
        else if(hours > 0) {
            builder.append(hours + "h ");
            builder.append(minutes + "m ");
            builder.append(seconds + "s ");
            builder.append(millis + "ms ");
        }
        else if(minutes > 0) {
            builder.append(minutes + "m ");
            builder.append(seconds + "s ");
            builder.append(millis + "ms ");
        }
        else if(seconds > 0) {
            builder.append(seconds + "s ");
            builder.append(millis + "ms ");
        }
        else {
            builder.append(millis + "ms ");
        }
        
        return builder.toString();
    }
    
    private double elaspsedTime(double timeMillis, int timeType){
        switch(timeType) {
            case TYPE_SECOND:
                return timeMillis/1000.00;
            case TYPE_MINUTE:
                return timeMillis/(1000.00 * 60);
            case TYPE_HOUR:
                return timeMillis/(1000.00 * 3600);
            case TYPE_DAY:
                return timeMillis/(1000.00 * 3600 * 24);
            default:
                return timeMillis;
        }
    }
}

