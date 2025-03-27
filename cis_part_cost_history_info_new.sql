-- ----------------------------
-- Table structure for table part_cost_history
-- ----------------------------
DROP TABLE IF EXISTS part_cost_history_new;
CREATE TABLE part_cost_history_new (
  `id` int NOT NULL AUTO_INCREMENT,
  `Part_Number` varchar(50) NOT NULL,
  `Cost_USD` decimal(10,5) NOT NULL,
  `Date` date NOT NULL,
  `created_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `modified_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`),
  UNIQUE KEY `unique_part_info` (`Part_Number`,`Date`),
  KEY `Date` (`Date`)
)  COMMENT='Stores part cost decimal (10,5) information at specific dates';

