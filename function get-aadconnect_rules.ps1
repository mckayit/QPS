function get-aadconnect_rules
{
    Import-Module adsync
    $ids = Get-ADSyncRule
    foreach ($ID in $IDs)
    {

        [PSCustomObject] @{
            Identifier       = $id.Identifier
            InternalId       = $id.InternalId
            Name             = $id.name
            "Internal_Id"    = $id.InternalId
            Description      = $id.Description
            #    ImmutableTag             = $id.ImmutableTag
            Connector        = $id.Connector
            Direction        = $id.Direction
            Disabled         = $id.Disabled
            SourceObjectType = $id.SourceObjectType | select-object 
            TargetObjectType = $id.TargetObjectType
            Precedence       = $id.Precedence
            PrecedenceAfter  = $id.PrecedenceAfter
            PrecedenceBefore = $id.PrecedenceBefore
            #   LinkType                 = $id.LinkType
            #   EnablePasswordSync       = $id.EnablePasswordSync
            JoinFilter       = ($id.JoinFilter | Select-Object  -ExpandProperty JoinConditionList )
            ScopeFilter      = $id.ScopeFilter  | Select-Object  -ExpandProperty ScopeConditionList
            # AttributeFlowMappings = $id.AttributeFlowMappings | Select-Object -ExpandProperty AttributeFlowMappings
            #   SoftDeleteExpiryInterval = $id.SoftDeleteExpiryInterval
            #   SourceNamespaceId        = $id.SourceNamespaceId
            #   TargetNamespaceId        = $id.TargetNamespaceId
            #   VersionAgnosticTag       = $id.VersionAgnosticTag
            #   TagVersion               = $id.TagVersion
            #   IsStandardRule           = $id.IsStandardRule
            #   IsLegacyCustomRule       = $id.IsLegacyCustomRule
            #   JoinHash                 = $id.JoinHash
        }
    }
}