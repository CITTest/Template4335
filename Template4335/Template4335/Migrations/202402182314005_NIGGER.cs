namespace Template4335.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class NIGGER : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.Orders", "RentTime", c => c.String());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.Orders", "RentTime", c => c.DateTime(nullable: false));
        }
    }
}
